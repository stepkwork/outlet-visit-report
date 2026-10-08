---
name: Visit record storage
description: Storage and sync expectations for outlet visit records.
---

Outlet visit drafts and saved records are stored locally in the browser. They do not sync across devices or users.

**Why:** The initial counter was scoped as a single-user field tool without accounts or a backend, so device-local persistence avoids implying shared or cloud storage.

**How to apply:** Preserve the local-only expectation unless the user asks for cross-device access, backup, or team sharing; explain the tradeoff before adding cloud persistence.
