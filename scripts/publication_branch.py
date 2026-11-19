"""Select the reviewed publication branch without changing legacy PC defaults."""
import os
from pathlib import Path
import subprocess

G10_BRANCH = 'g10-live-20261120'


def publication_branch():
    value = os.environ.get('BENTAN_PUBLICATION_BRANCH', 'main')
    if value not in ('main', G10_BRANCH):
        raise ValueError('Unreviewed publication branch')
    return value


def require_publication_branch(repo: Path):
    branch = publication_branch()
    if branch == 'main':
        return
    result = subprocess.run(['git', 'symbolic-ref', '--quiet', '--short', 'HEAD'],
                            cwd=repo, capture_output=True, text=True, timeout=15)
    if result.returncode or result.stdout.strip() != branch:
        raise ValueError('G10 publication requires its dedicated checked-out branch')
