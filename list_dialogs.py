import asyncio
import os

from dotenv import load_dotenv
from bale import BaleClient


load_dotenv()
TOKEN = os.environ.get("BALE_TOKEN", "").strip()


async def main():
    if not TOKEN:
        raise SystemExit("ابتدا python login.py را اجرا کن.")

    async with BaleClient(TOKEN) as client:
        me = await client.get_me()
        print(f"Logged in as: {me.title} | user:{me.peer.id}")
        print("-" * 90)

        dialogs = await client.get_dialogs(limit=300)
        for d in dialogs:
            kind = "user" if d.is_user else "channel"
            print(f"{kind}:{d.id}\t{d.title or '(no title)'}")


if __name__ == "__main__":
    asyncio.run(main())
