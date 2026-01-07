---
inclusion: always
---

## ManageBac Authentication Pattern

### Login Flow

The authentication system follows a standard session-based approach:

1. **Extract CSRF Token**: Retrieve the CSRF token from the login page meta tag (`csrf-token`)
2. **Authenticate**: Submit credentials to the ManageBac login endpoint
3. **Return Session Cookies**: Extract and return the session cookies from the response

### URL Structure

ManageBac uses school-specific subdomains:
- Format: `https://{school_code}.managebac.cn/`
- Example: `https://myschool.managebac.cn/login`

### Session Management

- Session cookies must be passed to all subsequent requests to maintain authentication
- Use the `cookies` parameter in `requests.get()` and similar methods
- Cookies are obtained from the login response and reused throughout the session

### Implementation Pattern

```python
# Login returns cookies
cookies = login(school_code, email, password)

# Cookies are passed to all authenticated requests
response = requests.get(task_url, cookies=cookies)
```

### Key Conventions

- All authenticated endpoints require the session cookies parameter
- Use BeautifulSoup with `lxml` parser for HTML parsing
- Include progress indicators (e.g., `tqdm.write()`) for user feedback during login