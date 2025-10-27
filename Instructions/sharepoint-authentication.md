# SharePoint Authentication Method

## Overview

The GALLOP Pilot Biopac analysis toolkit uses **OAuth2 Device Code Flow** for SharePoint authentication. This method provides secure, interactive authentication without requiring client secrets or passwords to be stored in the codebase.

## Authentication Method: Device Code Flow

### What is Device Code Flow?

Device Code Flow (also called Device Authorization Grant) is an OAuth2 authentication method designed for devices and applications that:
- Don't have a web browser
- Have limited input capabilities
- Run in command-line environments

This method is ideal for our Python scripts that need to access SharePoint data securely.

### How It Works

1. **Application requests authentication**
   - The script initiates a device code request to Microsoft's OAuth2 endpoint
   - Microsoft returns a device code and a user code

2. **User authenticates in browser**
   - User opens https://microsoft.com/devicelogin in their web browser
   - User enters the provided code
   - User signs in with their UF credentials (username/password + MFA if required)
   - User grants permissions to access SharePoint on their behalf

3. **Application receives access token**
   - After user approval, the application polls Microsoft's endpoint
   - Microsoft returns an access token
   - The script uses this token for all SharePoint API calls

4. **Token management**
   - Access tokens are valid for a limited time (typically 1 hour)
   - Refresh tokens can be used to obtain new access tokens
   - The `office365-sharepoint-python-client` library handles token management automatically

## Implementation Details

### Required Libraries

```python
from office365.sharepoint.client_context import ClientContext
```

We use the [`office365-sharepoint-python-client`](https://github.com/vgrem/Office365-REST-Python-Client) library, which provides a Python interface to SharePoint's REST API.

### Configuration Parameters

The authentication requires the following parameters (stored in `.env`):

| Parameter | Description | Example Value |
|-----------|-------------|---------------|
| `TENANT_ID` | Azure AD tenant identifier | `common` (multi-tenant) or `yourdomain.onmicrosoft.com` |
| `CLIENT_ID` | Application (client) ID | `d3590ed6-52b3-4102-aeff-aad2292ab01c` |
| `SHAREPOINT_SITE_URL` | SharePoint site base URL | `https://uflorida.sharepoint.com/sites/PRICE-GALLOP` |

### Client ID Used

We use Microsoft's public Office client ID: **`d3590ed6-52b3-4102-aeff-aad2292ab01c`**

This is a well-known Microsoft first-party client ID that:
- Does not require app registration in Azure AD
- Has built-in permissions for SharePoint access
- Works across Microsoft 365 tenants
- Requires no client secret (public client)

### Authentication Code

The authentication is implemented in `src/qst_logs/downloader.py:76-120`:

```python
def authenticate_interactive(self) -> bool:
    """
    Authenticate using device code flow.
    User will receive a code to enter at microsoft.com/devicelogin

    Returns:
        True if authentication successful, False otherwise
    """
    try:
        # Use device code flow
        self.ctx = ClientContext(self.site_url).with_device_flow(
            tenant=self.tenant_id,
            client_id=self.client_id
        )

        # Test connection
        web = self.ctx.web
        self.ctx.load(web)
        self.ctx.execute_query()

        print(f"\n[OK] Successfully connected to: {web.properties['Title']}\n")
        return True

    except Exception as e:
        print(f"\n[ERROR] Authentication failed: {str(e)}")
        return False
```

## User Experience

When running scripts that require SharePoint access, users see:

```
==================================================================
AUTHENTICATION REQUIRED
==================================================================

You will receive a CODE below.

STEPS TO AUTHENTICATE:
  1. Wait for the code to appear below
  2. Open your web browser and go to: https://microsoft.com/devicelogin
  3. Enter the code when prompted
  4. Sign in with your UF credentials
  5. Return to this terminal - it will continue automatically

==================================================================

To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code XXXXXXXXX to authenticate.
```

After successful authentication:

```
[OK] Successfully connected to: PRICE-GALLOP
```

## Security Considerations

### Advantages

1. **No stored credentials**: No passwords, secrets, or tokens stored in code or config files
2. **User authentication**: Each user authenticates with their own credentials
3. **MFA support**: Supports multi-factor authentication seamlessly
4. **Permission scoping**: Users can only access SharePoint resources they have permission to view
5. **Audit trail**: All actions are logged under the user's identity in SharePoint
6. **Token expiration**: Access tokens expire automatically, reducing risk of token theft

### Best Practices

1. **Never commit tokens**: The `.env` file is in `.gitignore` to prevent credential leakage
2. **User permissions**: Access is controlled by SharePoint permissions, not application logic
3. **Public client**: Using a public Microsoft client ID means no secrets to protect
4. **Interactive only**: This flow requires user interaction, preventing automated abuse

## Troubleshooting

### Common Issues

**Authentication fails with "invalid_client"**
- Verify `CLIENT_ID` in `.env` is correct
- Ensure `TENANT_ID` is set to `common` or your organization's tenant ID

**"Access Denied" errors**
- User doesn't have permission to the SharePoint site
- Contact SharePoint site administrator to request access

**Code expires before entry**
- Device codes typically expire after 15 minutes
- Re-run the script to get a new code

**Token refresh fails**
- Clear cached tokens (if library caches them)
- Re-authenticate by running the script again

## Alternative Authentication Methods

While we use Device Code Flow, the `office365-sharepoint-python-client` library supports other methods:

1. **User credentials** (username/password) - Not recommended, doesn't support MFA
2. **Client credentials** (client ID + secret) - Requires Azure AD app registration
3. **Certificate-based** - Requires certificate management
4. **Managed identity** - For Azure-hosted applications only

Device Code Flow was chosen because it:
- Works on local machines and servers
- Supports MFA and modern authentication
- Doesn't require Azure AD app registration
- Provides user-level authentication and auditing

## References

- [Microsoft Device Code Flow Documentation](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code)
- [office365-sharepoint-python-client Library](https://github.com/vgrem/Office365-REST-Python-Client)
- [OAuth 2.0 Device Authorization Grant](https://oauth.net/2/device-flow/)

## Related Files

- `src/qst_logs/downloader.py` - Main authentication and download implementation
- `download_qst_logs.py` - Script that uses the authentication
- `.env.example` - Configuration template
- `requirements.txt` - Includes `Office365-REST-Python-Client`
