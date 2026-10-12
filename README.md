# Outlook Supplier Claim Drafter

Python desktop application for preparing claim email drafts for outstanding supplier orders.

## Project Status

**Current phase:** Phase 00 — Repository Setup.

The functional application has not been implemented yet. Requirements, technical decisions, and phase status will be documented according to the master development protocol.

## Planned Objectives

- Retrieve outstanding order data from SQL Server in read-only mode.
- Allow users to review suppliers and outstanding order lines through a desktop interface.
- Prepare email drafts and Excel files for user review.
- Maintain user-specific local customizations and history.
- Keep development and corporate environment configurations separate.

The integration mechanism for modern Outlook and the history deduplication rules must be investigated and validated before implementation.

## Development Environment

Python 3.13 (64-bit) is the initially proposed version. Compatibility with the required dependencies and the corporate environment remains to be validated.

Virtual environment setup and execution instructions will be documented once they have been verified on the development machine.

## Project Structure

- `src/`: application source code, pending implementation.
- `tests/`: tests, pending implementation.
- `docs/`: master protocol, scope, decisions, and phase status.

## Security

- Do not store credentials in the repository.
- Do not commit local databases or supplier files.
- Do not write data to SQL Server.
- Do not send real emails during testing without explicit authorization.
- Generating a draft does not mean sending an email.

## Development Protocol

Consult `docs/protocolo-maestro.md` before making changes.