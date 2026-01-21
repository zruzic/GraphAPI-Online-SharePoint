# Security Policy

## Reporting a Vulnerability
If you find any security issues or accidentally committed secrets in this collection, please open an issue or contact me directly.

## Important Note
Always ensure you use environment variables (e.g., `{{ClientID}}` in Postman or `os.getenv()` in Python) and **never hardcode real secrets** in the collection files or scripts before committing changes.

For local testing, use a `.env` file or a local environment file that is excluded by the `.gitignore` included in this repository.
