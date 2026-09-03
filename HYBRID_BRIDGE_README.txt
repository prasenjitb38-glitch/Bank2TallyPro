Bank2Tally Hybrid Bridge

1. Keep TallyPrime open with the company loaded and HTTP Server enabled on port 9000.
2. Install Python 3.11+ on the Windows Tally computer and run:
   pip install -r requirements.txt
3. Create one strong random Bridge token. Add these Render environment variables:
   TALLY_BRIDGE_URL=https://YOUR-SECURE-TUNNEL.example
   TALLY_BRIDGE_TOKEN=the-same-token
4. Expose only local port 9010 through a private/secured tunnel. Do not expose port 9000 directly. For a temporary Cloudflare URL, install cloudflared and run:
   cloudflared tunnel --url http://127.0.0.1:9010
5. On the Tally computer, run Start_Tally_Bridge.bat and enter the same token.
6. Put the tunnel URL (without a trailing slash) in Render as TALLY_BRIDGE_URL and the same token as TALLY_BRIDGE_TOKEN.
7. Test https://YOUR-SECURE-TUNNEL.example/health before using Send to Tally.

You can use Start_Hybrid_Bridge.bat to start the local bridge and a temporary
Cloudflare tunnel together. The temporary URL changes whenever the tunnel is
restarted, so update TALLY_BRIDGE_URL in Render each time. For regular use,
create a named Cloudflare tunnel with a stable hostname.

The bridge forwards XML to local TallyPrime and performs scanned-PDF OCR with Windows OCR.
Never commit bank statements, databases, login files, or the bridge token to GitHub.
