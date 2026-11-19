"""One explicit API origin for existing publishers; missing config fails closed."""
import json
import os
from pathlib import Path
from urllib.parse import urlsplit

ROUTES = {'/api/recordings/publication', '/api/recordings/capture',
          '/api/journal-management/callback', '/api/journal-management/supplements'}


def checked_origin(value):
    if not isinstance(value, str):
        raise ValueError('Bentan API origin is required')
    parsed = urlsplit(value)
    if (parsed.scheme != 'https' or not parsed.hostname or parsed.username or parsed.password
            or parsed.path not in ('', '/') or parsed.query or parsed.fragment
            or any(char.isspace() for char in value)):
        raise ValueError('Bentan API origin must be an HTTPS origin')
    # Force validation of a malformed port before attaching publisher tokens.
    parsed.port
    return value.rstrip('/')


def api_base_url():
    override = os.environ.get('BENTAN_API_BASE_URL', '').strip()
    if override:
        return checked_origin(override)
    manifest = json.loads((Path(__file__).resolve().parent.parent / 'bentan-endpoints.json').read_text('utf-8'))
    if manifest.get('version') != 1:
        raise ValueError('Unsupported Bentan endpoint manifest')
    return checked_origin(manifest.get('apiBaseUrl'))


def endpoint(route):
    if route not in ROUTES:
        raise ValueError('Unknown Bentan integration route')
    return api_base_url() + route


def public_read_headers(value):
    parsed = urlsplit(value)
    approved = 'https://bentang10ops20261119.shares.zrok.io'
    if (checked_origin(api_base_url()) == approved and parsed.scheme == 'https'
            and parsed.hostname == 'bentang10ops20261119.shares.zrok.io'
            and parsed.port in (None, 443) and not parsed.username and not parsed.password
            and parsed.path in ('/api/recordings/publication', '/api/schedule/calendar')):
        return {'skip_zrok_interstitial': 'true'}
    return {}
