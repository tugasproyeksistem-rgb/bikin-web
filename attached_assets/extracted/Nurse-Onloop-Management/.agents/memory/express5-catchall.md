---
name: Express 5 catch-all route syntax
description: Correct wildcard route syntax for Express 5 with path-to-regexp v8
---

Express 5 uses path-to-regexp v8 which no longer supports bare `*` wildcards.

**Wrong:** `app.get("*", handler)` — throws `PathError: Missing parameter name at index 1: *`

**Correct:** `app.get("/{*splat}", handler)` — named splat parameter works fine

**Why:** path-to-regexp v8 requires named capture groups for wildcards.

**How to apply:** Any SPA fallback route, catch-all API 404 handler, or generic middleware that uses `*` as a path must use `/{*splat}` instead.
