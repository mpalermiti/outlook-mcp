"""Intent tools: request-level tools built on the classic ones (opt-in ``intents`` group).

The shared contracts live here, one module each: ``envelope`` (the answer shape),
``budget`` (request budget, deadline and limiter), ``resolver`` and ``people``
(names to things), ``selection`` (mail selections), ``ids`` (the id family) and
``transport`` (the counted raw Graph path). The design is in
``docs/design/intent-contracts.md``.
"""
