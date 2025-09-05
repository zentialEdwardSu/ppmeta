#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Simplified Changelog Generator
Quickly generate project changelog

Usage:
    python changelog.py              # Generate changelog for all commits
    python changelog.py --auto      # Auto continue from existing changelog
    python changelog.py --since abc123  # Start from specified commit
"""

import subprocess
import re
import os
import argparse
from datetime import datetime
from typing import List, Dict, Optional


class SimpleChangelogGenerator:
    """Simplified Changelog Generator"""

    # Supported commit types with corresponding emojis
    TYPES = {
        "feat": "✨ Features",
        "fix": "🐛 Bug Fixes",
        "docs": "📚 Documentation",
        "style": "💄 Styles",
        "refactor": "♻️ Code Refactoring",
        "perf": "⚡ Performance Improvements",
        "test": "✅ Tests",
        "build": "📦 Builds",
        "ci": "👷 Continuous Integration",
        "chore": "🔧 Chores",
        "revert": "⏪ Reverts",
    }

    def __init__(self):
        self.commit_pattern = re.compile(
            r"^(?P<type>\w+)(?:\((?P<scope>[^)]+)\))?(?P<breaking>!)?: (?P<subject>.+)$"
        )

    def get_commits(self, since_hash: Optional[str] = None) -> List[Dict]:
        """Get Git commits"""
        cmd = ["git", "log", "--pretty=format:%H|%an|%ad|%s", "--date=short"]

        if since_hash:
            cmd.append(f"{since_hash}..HEAD")

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            commits = []

            for line in result.stdout.strip().split("\n"):
                if not line:
                    continue

                parts = line.split("|", 3)
                if len(parts) < 4:
                    continue

                hash_val, author, date, message = parts
                commit = self.parse_commit(hash_val[:8], author, date, message)
                if commit:
                    commits.append(commit)

            return commits
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed to get Git commits: {e}")
            return []

    def parse_commit(self, hash_val: str, author: str, date: str, message: str) -> Dict:
        """Parse commit information"""
        match = self.commit_pattern.match(message)

        if match:
            groups = match.groupdict()
            return {
                "hash": hash_val,
                "author": author,
                "date": date,
                "type": groups["type"].lower(),
                "scope": groups["scope"] or "",
                "subject": groups["subject"],
                "breaking": groups["breaking"] == "!",
                "raw_message": message,
            }
        else:
            # Non-standard commits are categorized as chore
            return {
                "hash": hash_val,
                "author": author,
                "date": date,
                "type": "chore",
                "scope": "",
                "subject": message,
                "breaking": False,
                "raw_message": message,
            }

    def group_commits(self, commits: List[Dict]) -> Dict[str, List[Dict]]:
        """Group commits by type"""
        groups = {}
        breaking_changes = []

        for commit in commits:
            commit_type = commit["type"]

            if commit["breaking"]:
                breaking_changes.append(commit)

            if commit_type not in groups:
                groups[commit_type] = []
            groups[commit_type].append(commit)

        # Add breaking changes group
        if breaking_changes:
            groups["breaking"] = breaking_changes

        return groups

    def generate_changelog(
        self, commits: List[Dict], version: Optional[str] = None
    ) -> str:
        """Generate changelog content"""
        if not commits:
            return self.generate_empty_changelog()

        groups = self.group_commits(commits)
        latest_hash = commits[0]["hash"] if commits else ""

        # Generate content
        content = self.generate_header(version)
        content += self.generate_sections(groups)
        content += self.generate_footer(commits, latest_hash)

        return content

    def generate_header(self, version: Optional[str] = None) -> str:
        """Generate header"""
        header = "# Changelog\n\n"

        if version:
            header += f"## [{version}] - {datetime.now().strftime('%Y-%m-%d')}\n\n"
        else:
            header += f"## [Unreleased] - {datetime.now().strftime('%Y-%m-%d')}\n\n"

        return header

    def generate_sections(self, groups: Dict[str, List[Dict]]) -> str:
        """Generate sections"""
        content = ""

        # Sort: breaking changes first, then by importance
        type_order = ["breaking"] + list(self.TYPES.keys())

        for commit_type in type_order:
            if commit_type not in groups:
                continue

            commits = groups[commit_type]
            if not commits:
                continue

            # Section title
            if commit_type == "breaking":
                title = "💥 Breaking Changes"
            else:
                title = self.TYPES.get(commit_type, f"🔄 {commit_type.upper()}")

            content += f"### {title}\n\n"

            # Commit list
            for commit in commits:
                content += self.format_commit(commit)

            content += "\n"

        return content

    def format_commit(self, commit: Dict) -> str:
        """Format commit"""
        line = "- "

        # Add scope
        if commit["scope"]:
            line += f"**{commit['scope']}**: "

        # Add subject
        line += commit["subject"]

        # Breaking change marker
        if commit["breaking"]:
            line += " ⚠️"

        # Add hash and author
        line += f" ([{commit['hash']}]) by {commit['author']}\n"

        return line

    def generate_footer(self, commits: List[Dict], latest_hash: str) -> str:
        """Generate footer with latest commit hash at the bottom"""
        footer = "\n---\n\n**Summary**\n"
        footer += f"- Commits in this update: {len(commits)}\n"
        footer += f"- Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"

        if latest_hash:
            footer += "**Next generation usage**:\n"
            footer += f"```bash\npython changelog.py --since {latest_hash}\n```\n\n"

        # Add the latest commit hash at the very bottom for easy retrieval
        if latest_hash:
            footer += f"<!-- LATEST_COMMIT_HASH: {latest_hash} -->\n"

        return footer

    def generate_empty_changelog(self) -> str:
        """Generate empty changelog"""
        content = "# Changelog\n\n"
        content += "This document records important changes to the project.\n\n"
        content += f"## [Unreleased] - {datetime.now().strftime('%Y-%m-%d')}\n\n"
        content += "No changes recorded yet.\n"
        return content

    def get_last_commit_from_changelog(self, file_path: str) -> Optional[str]:
        """Get the last commit hash from existing changelog"""
        if not os.path.exists(file_path):
            return None

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()

            # Look for the latest commit hash in the comment at the end
            pattern = r"<!-- LATEST_COMMIT_HASH: ([a-f0-9]{8}) -->"
            match = re.search(pattern, content)
            return match.group(1) if match else None
        except Exception:
            return None

    def get_latest_commit_hash(self) -> str:
        """Get the latest commit hash"""
        try:
            result = subprocess.run(
                ["git", "rev-parse", "HEAD"], capture_output=True, text=True, check=True
            )
            return result.stdout.strip()[:8]
        except Exception:
            return ""


def main():
    parser = argparse.ArgumentParser(description="Generate project changelog")
    parser.add_argument("--since", help="Starting commit hash")
    parser.add_argument("--version", help="Version number")
    parser.add_argument("--output", default="CHANGELOG.md", help="Output file")
    parser.add_argument(
        "--auto", action="store_true", help="Auto continue from existing changelog"
    )

    args = parser.parse_args()

    generator = SimpleChangelogGenerator()

    # Determine starting commit
    since_hash = args.since
    if args.auto and not since_hash:
        since_hash = generator.get_last_commit_from_changelog(args.output)
        if since_hash:
            print(f"🔍 Auto-detected last processed commit: {since_hash}")

    # Get commits
    print("📝 Getting Git commits...")
    commits = generator.get_commits(since_hash)

    if not commits:
        print("ℹ️  No new commits found")
        return

    print(f"📊 Found {len(commits)} new commits")

    # Generate changelog
    print("🔄 Generating changelog...")
    changelog = generator.generate_changelog(commits, args.version)

    # Write to file
    try:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(changelog)

        print(f"✅ Changelog generated: {args.output}")

        # Show latest commit hash for next use
        latest_hash = generator.get_latest_commit_hash()
        if latest_hash:
            print(f"📌 Latest commit: {latest_hash}")
            print(f"💡 Next usage: python changelog.py --since {latest_hash}")

    except Exception as e:
        print(f"❌ Failed to write file: {e}")


if __name__ == "__main__":
    main()
