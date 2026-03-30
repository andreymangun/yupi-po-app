import os
from dotenv import load_dotenv

try:
    from supabase import Client, create_client
except Exception:  # pragma: no cover
    Client = object
    create_client = None

load_dotenv()


def get_supabase() -> Client:
    url = os.getenv("SUPABASE_URL")
    key = os.getenv("SUPABASE_KEY")
    if not url or not key:
        raise ValueError("SUPABASE_URL atau SUPABASE_KEY belum diset.")
    if create_client is None:
        raise ImportError("Package supabase belum terpasang.")
    return create_client(url, key)
