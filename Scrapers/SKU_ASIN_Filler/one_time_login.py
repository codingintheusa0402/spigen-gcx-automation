"""
Run this ONCE, interactively, to log into the 3 Seller Central accounts
inside a dedicated persistent browser profile. Once logged in, the cookies
persist in PROFILE_DIR and fill_asin_blanks.py (run unattended via launchd)
reuses them without needing to log in again.

Usage: python3 one_time_login.py
Opens a visible Chrome window. Log into each tab manually (including any
2FA/OTP), then press Enter in this terminal when all 3 are logged in.
"""
import asyncio
import os
from playwright.async_api import async_playwright

PROFILE_DIR = os.path.expanduser("~/.chrome-sc-inventory-profile")

URLS = [
    "https://sellercentral.amazon.de/myinventory/inventory",
    "https://sellercentral-japan.amazon.com/amazonsell/fba-inventory",
    "https://sellercentral.amazon.com/inventoryplanning/manageinventoryhealth",
]


async def main():
    async with async_playwright() as pw:
        ctx = await pw.chromium.launch_persistent_context(
            PROFILE_DIR, channel="chrome", headless=False
        )
        pages = []
        first = ctx.pages[0] if ctx.pages else await ctx.new_page()
        await first.goto(URLS[0])
        pages.append(first)
        for url in URLS[1:]:
            p = await ctx.new_page()
            await p.goto(url)
            pages.append(p)

        print("\n3 tabs opened: DE inventory, JP FBA inventory, US inventory health.")
        print("Log into each one (2FA/OTP as needed) until each shows the real")
        print("inventory page (not a sign-in screen). Cookies save to disk live,")
        print("as you go -- no need to wait for this script to exit.")
        print("This process will stay open; stop it (Ctrl+C, or have it killed)")
        print("once all 3 tabs are logged in.")
        await asyncio.Event().wait()


if __name__ == "__main__":
    asyncio.run(main())
