#!/usr/bin/env python3
"""
RNS Article Manager — Add new articles to the RNS feed and push to GitHub.

Usage:
  python add-rns.py                          # Interactive mode
  python add-rns.py --file new_articles.json # Bulk import from JSON file
  python add-rns.py --list                   # List all companies in the database
  python add-rns.py --count                  # Show article count

The JSON file for --file should be an array of objects:
[
  {
    "company": "Company PLC",
    "publish_date": "07 April 2026",
    "time_period": "Q1 2026 results",
    "details": [
      "First bullet point",
      "Second bullet point"
    ]
  }
]
"""

import json, os, subprocess, sys, argparse
from datetime import datetime

REPO_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(REPO_DIR, "rns-data", "rns-articles.json")

def load_articles():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r') as f:
            return json.load(f)
    return []

def save_articles(articles):
    os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
    with open(DATA_FILE, 'w') as f:
        json.dump(articles, f, indent=2)
    print(f"  Saved {len(articles)} total articles to {DATA_FILE}")

def git_push(message):
    os.chdir(REPO_DIR)
    try:
        subprocess.run(["git", "add", "-A"], check=True, capture_output=True)
        result = subprocess.run(["git", "status", "--porcelain"], capture_output=True, text=True)
        if not result.stdout.strip():
            print("  No changes to push.")
            return
        subprocess.run(["git", "commit", "-m", message], check=True, capture_output=True)
        subprocess.run(["git", "push"], check=True, capture_output=True)
        print(f"  Pushed to GitHub: {message}")
    except subprocess.CalledProcessError as e:
        print(f"  Git error: {e.stderr.decode() if e.stderr else e}")
        print("  You may need to push manually: cd equity-blog && git push")

def add_interactive():
    articles = load_articles()
    print("\n=== Add New RNS Article ===\n")
    company = input("Company name: ").strip()
    if not company:
        print("Cancelled.")
        return
    pub_date = input("Publish date (e.g. 07 April 2026): ").strip()
    time_period = input("Time period: ").strip()
    print("Enter bullet points (one per line, empty line to finish):")
    details = []
    while True:
        line = input("  • ").strip()
        if not line:
            break
        details.append(line)

    if not details:
        print("No details entered. Cancelled.")
        return

    article = {
        "company": company,
        "publish_date": pub_date,
        "time_period": time_period,
        "details": details
    }
    articles.append(article)
    save_articles(articles)
    git_push(f"RNS: {company} — {pub_date}")
    print("Done!\n")

def add_from_file(filepath):
    articles = load_articles()
    with open(filepath, 'r') as f:
        new_articles = json.load(f)
    if not isinstance(new_articles, list):
        new_articles = [new_articles]
    count = len(new_articles)
    articles.extend(new_articles)
    save_articles(articles)
    names = ", ".join(set(a["company"] for a in new_articles))
    git_push(f"RNS: Added {count} articles ({names})")
    print(f"Added {count} articles.\n")

def list_companies():
    articles = load_articles()
    companies = sorted(set(a["company"] for a in articles))
    print(f"\n{len(companies)} companies in database:\n")
    for c in companies:
        count = sum(1 for a in articles if a["company"] == c)
        print(f"  {c} ({count})")
    print()

def show_count():
    articles = load_articles()
    print(f"\n{len(articles)} articles in database.\n")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="RNS Article Manager")
    parser.add_argument("--file", help="JSON file of new articles to import")
    parser.add_argument("--list", action="store_true", help="List all companies")
    parser.add_argument("--count", action="store_true", help="Show article count")
    args = parser.parse_args()

    if args.list:
        list_companies()
    elif args.count:
        show_count()
    elif args.file:
        add_from_file(args.file)
    else:
        add_interactive()
