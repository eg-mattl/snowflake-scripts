# Snowflake user default-role export

`export_user_default_roles.py` reads Snowflake account metadata and writes an Excel workbook with each non-deleted user and the account's configured default role.

## Scope and caveats

- Source: `SNOWFLAKE.ACCOUNT_USAGE.USERS`.
- Filter: `DELETED_ON IS NULL`. Disabled accounts and `SNOWFLAKE_SERVICE` users are included; only deleted users are excluded.
- A SQL `NULL` in `DEFAULT_ROLE` is shown as **`(not set)`** in Excel.
- `ACCOUNT_USAGE` can lag recent account changes by about two hours.
- The executing role must be able to read `SNOWFLAKE.ACCOUNT_USAGE.USERS`.
- The Summary worksheet records the Snowflake session context, extraction time (UTC), row counts, source, scope, and freshness caveat. The filterable Users worksheet includes user name, login name, display name, type, disabled flag, and default role.

## Install

Use a virtual environment, then install the project dependencies:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Python 3.10 or newer is recommended.

## Connect with browser SSO

Set the account and user for the Snowflake Python Connector. The default authenticator is `externalbrowser`:

```bash
export SNOWFLAKE_ACCOUNT="your-account-identifier"
export SNOWFLAKE_USER="your-user"
export SNOWFLAKE_AUTHENTICATOR="externalbrowser"
# Optional role override; it must be authorized to read ACCOUNT_USAGE.USERS.
export SNOWFLAKE_ROLE="your-read-only-role"
# Supply a warehouse when the user has no default warehouse; SELECT needs one.
export SNOWFLAKE_WAREHOUSE="your-warehouse"

python export_user_default_roles.py --output ./reports/snowflake_user_default_roles.xlsx
```

The connector will open the browser-based authentication flow. No password is embedded in the script or passed on the command line.

## Connect with key-pair authentication

For an authorized key-pair connection, use the Snowflake Connector's JWT authenticator and point to a private key stored outside the repository:

```bash
export SNOWFLAKE_ACCOUNT="your-account-identifier"
export SNOWFLAKE_USER="your-service-user"
export SNOWFLAKE_AUTHENTICATOR="SNOWFLAKE_JWT"
export SNOWFLAKE_PRIVATE_KEY_FILE="/secure/path/to/private-key.p8"
# If the private key is encrypted, provide its passphrase via a secure environment:
export SNOWFLAKE_PRIVATE_KEY_FILE_PWD="..."

python export_user_default_roles.py
```

Do not commit private keys, passwords, tokens, or local `.env` files.

## Output behavior

By default, the workbook is written to `./snowflake_user_default_roles.xlsx`. Existing output files are not overwritten unless `--force` is supplied:

```bash
python export_user_default_roles.py --output ./reports/users.xlsx --force
```

The script reports connector and file errors to stderr and exits non-zero. It reads metadata only; it does not modify Snowflake objects or user settings.

## Tests

The unit tests use a fake connector and do not connect to Snowflake:

```bash
python -m unittest discover -s tests -v
```
