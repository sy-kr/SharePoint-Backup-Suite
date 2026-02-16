Place your .pfx certificate file here as `spbackup.pfx` for automatic discovery.

The certificate is required only for downloading Microsoft List attachments,
which uses the SharePoint REST API with certificate-based authentication.

If you don't need list attachments, no certificate is required.

To set up:
1. Generate or obtain an X.509 certificate (.pfx with private key)
2. Upload the public key (.cer) to your Azure AD app registration
3. Place the .pfx file here as `spbackup.pfx`
4. Optionally set CERT_PATH env var to override the auto-discovery path
5. Optionally set CERT_PASSWORD env var if the .pfx is password-protected
