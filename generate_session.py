# -*- coding: utf-8 -*-
"""
Helper script to generate a Telethon StringSession token.
This token lets Vercel connect to your Telegram account statelessly via an environment variable.
"""

import asyncio
import os
from telethon import TelegramClient
from telethon.sessions import StringSession

API_ID   = int(os.environ.get("TELEGRAM_API_ID", 0))
API_HASH = os.environ.get("TELEGRAM_API_HASH", "YOUR_API_HASH")

async def main():
    # Load existing session file
    client = TelegramClient('session', API_ID, API_HASH)
    await client.connect()
    
    if not await client.is_user_authorized():
        print("Error: session file is not authorized. Please run find_group.py first.")
        return
        
    string_session = StringSession.save(client.session)
    print("\n" + "=" * 80)
    print("YOUR TELEGRAM_SESSION_STRING (Copy this string to Vercel Environment Variables):")
    print("=" * 80)
    print(string_session)
    print("=" * 80 + "\n")

if __name__ == '__main__':
    asyncio.run(main())
