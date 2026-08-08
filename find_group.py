"""
Helper script to list all Telegram groups and channels the user belongs to.
Displays the title, ID/username, and type of each group or channel.
"""

import sys
from telethon import TelegramClient
from telethon.errors import RPCError

import os

API_ID   = int(os.environ.get("TELEGRAM_API_ID", 0))       # Set TELEGRAM_API_ID env var or paste your API ID
API_HASH = os.environ.get("TELEGRAM_API_HASH", "YOUR_API_HASH")  # Set TELEGRAM_API_HASH env var or paste your API Hash

client = TelegramClient('session', API_ID, API_HASH)


async def main() -> None:
    """Iterate through Telegram dialogs and display groups and channels."""
    print("Fetching Telegram groups and channels...\n")
    async for dialog in client.iter_dialogs():
        if dialog.is_group or dialog.is_channel:
            dialog_type = "Channel" if (dialog.is_channel and not dialog.is_group) else "Group"
            username = getattr(dialog.entity, 'username', None)
            identifier = f"@{username}" if username else str(dialog.id)
            print(f"Title: {dialog.name} | ID: {identifier} | Type: {dialog_type}")

    print("\nPlease copy the target Group/Channel ID or @username into telegram_to_excel.py.")


if __name__ == '__main__':
    try:
        with client:
            client.loop.run_until_complete(main())
    except RPCError as e:
        print(f"Telegram API Error: {e}", file=sys.stderr)
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
    except Exception as e:
        print(f"An error occurred: {e}", file=sys.stderr)
