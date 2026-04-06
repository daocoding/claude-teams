# Anthropic Official Plugin Validation Recommendations

This repository is close to a strong pre-approval plugin, but there are a few high-impact items to address before submitting for official validation.

## Critical (should fix before submission)

1. **Webhook authenticity is not validated (JWT verification TODO)**
   - Current webhook handler accepts any POST payload to `TEAMS_WEBHOOK_PATH`.
   - Add Bot Framework JWT verification for inbound requests and reject invalid tokens with 401.
   - Keep network-level protections, but plugin validation typically expects application-layer authentication for webhook sources.

2. **DM policy enforcement mismatch**
   - `allowlist` mode should silently drop unknown DMs instead of issuing pairing codes.
   - This is now fixed in `gate(...)` so behavior matches `ACCESS.md` and `README.md`.

## High Priority (recommended for validation robustness)

3. **Input validation for tool arguments**
   - Add stricter validation for `conversation_id`, `activity_id`, and `url`.
   - Reject empty strings and malformed URLs with clear errors.

4. **Attachment filename sanitization**
   - Current filename comes from remote metadata.
   - Strip path separators and control characters before writing to disk.

5. **Rate limiting and replay protection on webhook endpoint**
   - Add basic request throttling and replay-window checks after JWT validation to reduce abuse risk.

6. **Versioning and release metadata**
   - Keep plugin `version` aligned across package metadata and MCP server metadata.
   - Add a changelog for validator review and reproducibility.

## Medium Priority (quality / operability)

7. **Automated checks in CI**
   - Add at least: formatting, TypeScript check, and a smoke test for startup.
   - Include tests for DM policy behavior (`pairing` vs `allowlist` vs `disabled`).

8. **Document secure deployment defaults**
   - Keep single-tenant guidance prominent.
   - Provide explicit examples for reverse proxy auth and Bot Framework IP restriction.

9. **Structured logging for validator diagnostics**
   - Emit JSON logs with event types (`webhook_received`, `gate_drop`, `tool_reply`, etc.) for easier traceability.

## Already Improved in This Branch

- Added `.env.example` referenced by the README to improve setup completeness.
- Updated DM gate logic to honor `allowlist` and `disabled` behavior for DMs.
- Removed accidental `senderMap` overload that stored `serviceUrl` under synthetic keys.
