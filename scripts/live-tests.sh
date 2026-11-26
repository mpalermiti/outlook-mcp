#!/usr/bin/env bash
# Run outlook-mcp's live test tiers in a container, against the project's test accounts.
#
#   scripts/live-tests.sh [options] TIER [-- extra pytest arguments]
#
# TIER
#   live         read-only query-shape guards                       (RELEASING.md 1b)
#   integration  read-only response-shape smoke tests                (1c)
#   live_write   creates and deletes marked events, contacts, tasks
#                and unsent drafts on the main test account          (1d)
#   live_send    real mail and invitations from the main test account
#                to its pair, and nobody else; needs --sends N       (1e)
#   preflight    scripts/preflight.py, the read-only endpoint sweep  (1a)
#   leftovers    lists any marked item a run left on either account (read-only)
#   clean        deletes those, where deleting one sends nothing (a meeting with
#                attendees is listed, not deleted), then lists what is left
#   status       `outlook-mcp status` for both accounts (read-only)
#   offline      the offline suite, with no network and no account: the way to
#                run it on another Python (--python 3.10)
#
# Options
#   --env-file PATH   the run's env file (default: $OUTLOOK_MCP_LIVE_ENV_FILE)
#   --python X.Y      the image's Python (default 3.13)
#   --sends N         live_send only: the most real sends this run may make (at most 10)
#
# The env file lives OUTSIDE this repository, and nothing committed names the
# accounts. It sets:
#   OUTLOOK_MCP_TEST_ACCOUNT=<the main test account's address>
#   OUTLOOK_MCP_TEST_PAIR=<its pair's address>
#   OUTLOOK_MCP_LIVE_VOLUME=<Docker volume holding the main account's settings>
#   OUTLOOK_MCP_LIVE_PAIR_VOLUME=<the same for the pair>
#   OUTLOOK_MCP_TEST_PAIR_CAN_SEND=1   optional: once the pair account may send
#   OUTLOOK_MCP_LIVE_STOP_FILE=<absolute path>   optional: while a file exists
#       there, every tier but `offline` refuses to start. Whatever else uses the
#       accounts creates it when Microsoft locks or blocks one; a person deletes
#       it once the account is unlocked
#
# Each account is signed in once, by hand, inside a container and never on the
# host. Build the image first (any tier does, e.g. `offline`), create the volume,
# put a config.json in it with the app's client_id, "tenant_id": "consumers" and
# "allow_unencrypted_token_cache": true (the slim image has no keyring), then:
#   docker run --rm -it -v <volume>:/config outlook-mcp-live:py3.13 outlook-mcp auth
# RELEASING.md §1 has the whole setup, and what to run without the accounts
# (the offline suite, and optionally the own-mailbox `live` check).
#
# Needs Docker (in Claude Code, run it with the sandbox off).
set -euo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd -P)"

usage() { sed -n '2,/^set -euo/p' "$0" | sed '$d' | sed -E 's/^# ?//'; }
die() {
  echo "live-tests: $*" >&2
  exit 2
}

env_file="${OUTLOOK_MCP_LIVE_ENV_FILE:-}"
python_version="3.13"
sends=""

while [ $# -gt 0 ]; do
  case "$1" in
    --env-file)
      [ $# -ge 2 ] || die "--env-file needs a path"
      env_file="$2"
      shift 2
      ;;
    --python)
      [ $# -ge 2 ] || die "--python needs a version"
      python_version="$2"
      shift 2
      ;;
    --sends)
      [ $# -ge 2 ] || die "--sends needs a count"
      sends="$2"
      shift 2
      ;;
    -h | --help)
      usage
      exit 0
      ;;
    -*) die "unknown option: $1 (see --help)" ;;
    *) break ;;
  esac
done

tier="${1:-}"
if [ -z "$tier" ]; then
  usage
  exit 2
fi
shift
if [ "${1:-}" = "--" ]; then shift; fi

case "$python_version" in
  3.[0-9] | 3.[0-9][0-9]) ;;
  *) die "--python takes a version like 3.12" ;;
esac
tag="outlook-mcp-live:py${python_version}"
image=""

# BuildKit, because only it reads the Dockerfile's own ignore file (the build
# would still copy by name without it). The run uses the id this build produced,
# not the tag, so two checkouts building at once cannot run each other's image.
build() {
  echo "live-tests: building $tag from the working tree" >&2
  local idfile
  idfile="$(mktemp)"
  DOCKER_BUILDKIT=1 docker build --quiet --iidfile "$idfile" \
    --file "$REPO/docker/live-tests.Dockerfile" \
    --build-arg "PYTHON_VERSION=$python_version" --tag "$tag" "$REPO" >/dev/null
  image="$(cat "$idfile")"
  rm -f "$idfile"
  [ -n "$image" ] || die "the build produced no image"
}

if [ "$tier" = "offline" ]; then
  build
  exec docker run --rm --network none -e HOME=/tmp "$image" \
    python -m pytest -p no:cacheprovider -q "$@"
fi

# ── The env file: outside the repository, and complete ────────────────────

[ -n "$env_file" ] || die "no env file: pass --env-file PATH or set OUTLOOK_MCP_LIVE_ENV_FILE"
[ -f "$env_file" ] && [ -r "$env_file" ] || die "the env file is missing or unreadable"
# Where the file really is, through any symlinks, so a link from outside cannot
# point back into the repository.
real_file="$env_file"
for _ in 1 2 3 4 5 6 7 8; do
  [ -L "$real_file" ] || break
  link="$(readlink "$real_file")"
  case "$link" in
    /*) real_file="$link" ;;
    *) real_file="$(dirname "$real_file")/$link" ;;
  esac
done
[ -L "$real_file" ] && die "the env file is a chain of links too long to follow"
env_dir="$(cd "$(dirname "$real_file")" && pwd -P)"
# A worktree lives inside the main checkout: check that too.
common_git="$(git -C "$REPO" rev-parse --path-format=absolute --git-common-dir 2>/dev/null || true)"
main_repo="$REPO"
if [ -n "$common_git" ]; then main_repo="$(cd "$(dirname "$common_git")" && pwd -P)"; fi
for root in "$REPO" "$main_repo"; do
  case "$env_dir/" in
    "$root"/*) die "the env file is inside this repository; keep it outside it" ;;
  esac
done

# A value from the env file. Read for the volume names only, and never printed.
value_of() { sed -n "s/^$1=//p" "$env_file" | tail -n 1; }

for name in OUTLOOK_MCP_TEST_ACCOUNT OUTLOOK_MCP_TEST_PAIR \
  OUTLOOK_MCP_LIVE_VOLUME OUTLOOK_MCP_LIVE_PAIR_VOLUME; do
  [ -n "$(value_of "$name")" ] || die "the env file does not set $name"
done

# A locked or blocked account gets no more traffic, not even a sign-in check:
# Microsoft treats each attempt as more of what got it locked.
stop_file="$(value_of OUTLOOK_MCP_LIVE_STOP_FILE)"
if [ -n "$stop_file" ]; then
  case "$stop_file" in
    /*) ;;
    *) die "OUTLOOK_MCP_LIVE_STOP_FILE must be an absolute path" ;;
  esac
  if [ -e "$stop_file" ]; then
    die "the stop file OUTLOOK_MCP_LIVE_STOP_FILE names exists: a test account is locked or blocked. Unlock it at outlook.com in a browser, then delete the stop file"
  fi
fi
volume="$(value_of OUTLOOK_MCP_LIVE_VOLUME)"
pair_volume="$(value_of OUTLOOK_MCP_LIVE_PAIR_VOLUME)"
for name in OUTLOOK_MCP_LIVE_VOLUME OUTLOOK_MCP_LIVE_PAIR_VOLUME; do
  # `docker run -v` would quietly create an empty volume in its place.
  docker volume inspect "$(value_of "$name")" >/dev/null 2>&1 ||
    die "the Docker volume $name names does not exist"
done

# ── What to run ───────────────────────────────────────────────────────────

write=0
send=0
budget=0
check_after=0
pytest_run=(python -m pytest -p no:cacheprovider -v -rs)
case "$tier" in
  live | integration) cmd=("${pytest_run[@]}" -m "$tier") ;;
  live_write)
    write=1
    check_after=1
    cmd=("${pytest_run[@]}" -m live_write)
    ;;
  live_send)
    case "$sends" in
      '' | *[!0-9]*) die "live_send needs --sends N: the most real sends this run may make" ;;
    esac
    # New accounts have tiny sending limits, and Outlook.com locks one that sends
    # too much too soon (tests/live_guard.MAX_SEND_BUDGET holds the same line).
    [ "$sends" -le 10 ] || die "--sends is at most 10"
    send=1
    budget="$sends"
    check_after=1
    cmd=("${pytest_run[@]}" -m live_send)
    ;;
  preflight) cmd=(python scripts/preflight.py) ;;
  leftovers) cmd=(python -m tests.live_mailbox check-leftovers) ;;
  clean) cmd=(python -m tests.live_mailbox check-leftovers --clean) ;;
  status)
    cmd=(sh -c 'echo "main:"; outlook-mcp status; echo; echo "pair:";
      HOME=/pair OUTLOOK_MCP_CONFIG_DIR=/pair outlook-mcp status')
    ;;
  *) die "unknown tier: $tier (see --help)" ;;
esac
case "$tier" in
  live | integration | live_write | live_send) ;;
  *) set -- ;; # extra arguments are for pytest
esac

# The env file comes first, so the flags after it win: the own-mailbox opt-in
# never applies in here, and only this tier's write or send switch is on.
run=(docker run --rm --env-file "$env_file"
  -v "$volume:/config"
  -e OUTLOOK_MCP_LIVE_OWN_MAILBOX=0
  -e "OUTLOOK_MCP_LIVE_WRITE=$write"
  -e "OUTLOOK_MCP_LIVE_SEND=$send"
  -e "OUTLOOK_MCP_LIVE_SEND_BUDGET=$budget")
# The pair's sign-in goes only where the pair is used.
case "$tier" in
  live_send | leftovers | clean | status)
    run+=(-v "$pair_volume:/pair" -e OUTLOOK_MCP_TEST_PAIR_CONFIG_DIR=/pair)
    ;;
  *) run+=(-e OUTLOOK_MCP_TEST_PAIR_CONFIG_DIR=) ;;
esac

build
status=0
"${run[@]}" "$image" "${cmd[@]}" "$@" || status=$?
if [ "$check_after" = 1 ]; then
  echo "live-tests: checking for anything the run left behind" >&2
  "${run[@]}" "$image" python -m tests.live_mailbox check-leftovers || status=$((status ? status : 1))
fi
exit "$status"
