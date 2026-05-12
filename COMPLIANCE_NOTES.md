# Compliance Notes

This app is a customer recovery and data cleanup utility. It should only process files that the operator/client is legally authorized to access and provide.

## Hard boundaries

- Do not add marketplace buyers to marketing lists without appropriate consent or another lawful basis.
- Do not scrape Etsy, buyer profiles, private messages, or restricted account areas.
- Do not bypass platform restrictions.
- Do not include a bulk-email sender in the MVP.
- Keep a do-not-contact/suppression export available.

## Operator language

Use: "Recover and organize authorized customer records."

Avoid: "Harvest emails."

## Recommended export policy

- Marketing Eligible: only confirmed opt-in or legally reviewed basis.
- Transactional Only: order/support/admin communication only.
- Needs Review: unresolved consent or data quality issue.
- Do Not Contact: opt-out, invalid, suppressed, or otherwise risky.
