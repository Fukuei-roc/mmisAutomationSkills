from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path


REPO_ROOT = Path(r"C:\Users\NMMIS\.codex\skills")
REMOTE_NAME = "origin"
REMOTE_URL = "https://github.com/Fukuei-roc/mmisAutomationSkills.git"
GIT_USER_NAME = "Fukuei-roc"
GIT_USER_EMAIL = "f113097@yahoo.com.tw"


class GitPublisherError(RuntimeError):
    pass


def run_git(*args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    completed = subprocess.run(
        ["git", "-C", str(REPO_ROOT), *args],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    if check and completed.returncode != 0:
        raise GitPublisherError(completed.stderr.strip() or completed.stdout.strip() or "git command failed")
    return completed


def is_git_repo() -> bool:
    completed = run_git("rev-parse", "--is-inside-work-tree", check=False)
    return completed.returncode == 0 and completed.stdout.strip() == "true"


def ensure_repo_initialized() -> bool:
    if is_git_repo():
        return False
    subprocess.run(
        ["git", "init", str(REPO_ROOT)],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    return True


def ensure_local_config() -> None:
    run_git("config", "user.name", GIT_USER_NAME)
    run_git("config", "user.email", GIT_USER_EMAIL)


def get_remote_url(name: str) -> str | None:
    completed = run_git("remote", "get-url", name, check=False)
    if completed.returncode != 0:
        return None
    return completed.stdout.strip() or None


def ensure_remote() -> dict[str, str | bool | None]:
    current = get_remote_url(REMOTE_NAME)
    if current is None:
        run_git("remote", "add", REMOTE_NAME, REMOTE_URL)
        return {"remote_created": True, "remote_updated": False, "remote_url": REMOTE_URL}
    if current != REMOTE_URL:
        run_git("remote", "set-url", REMOTE_NAME, REMOTE_URL)
        return {"remote_created": False, "remote_updated": True, "remote_url": REMOTE_URL}
    return {"remote_created": False, "remote_updated": False, "remote_url": current}


def working_tree_changes() -> list[str]:
    completed = run_git("status", "--short", check=False)
    if completed.returncode != 0:
        raise GitPublisherError(completed.stderr.strip() or "unable to read git status")
    return [line for line in completed.stdout.splitlines() if line.strip()]


def current_branch() -> str:
    completed = run_git("branch", "--show-current", check=False)
    branch = completed.stdout.strip()
    if branch:
        return branch

    completed = run_git("symbolic-ref", "--short", "HEAD", check=False)
    branch = completed.stdout.strip()
    return branch or "main"


def ensure_branch(name: str) -> str:
    branch = current_branch()
    if branch:
        return branch
    run_git("checkout", "-b", name)
    return name


def commit_and_push(message: str, skip_push: bool) -> dict[str, object]:
    changes = working_tree_changes()
    if not changes:
        return {
            "has_changes": False,
            "committed": False,
            "pushed": False,
            "status_lines": [],
            "commit_message": message,
        }

    run_git("add", ".")

    commit = run_git("commit", "-m", message, check=False)
    if commit.returncode != 0:
        stderr = commit.stderr.strip()
        stdout = commit.stdout.strip()
        if "nothing to commit" in stderr.lower() or "nothing to commit" in stdout.lower():
            return {
                "has_changes": False,
                "committed": False,
                "pushed": False,
                "status_lines": changes,
                "commit_message": message,
            }
        raise GitPublisherError(stderr or stdout or "git commit failed")

    branch = ensure_branch("main")
    result: dict[str, object] = {
        "has_changes": True,
        "committed": True,
        "pushed": False,
        "status_lines": changes,
        "commit_message": message,
        "branch": branch,
    }
    if skip_push:
        return result

    push = run_git("push", "-u", REMOTE_NAME, branch, check=False)
    if push.returncode != 0:
        raise GitPublisherError(push.stderr.strip() or push.stdout.strip() or "git push failed")
    result["pushed"] = True
    return result


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Manage and publish C:\\Users\\NMMIS\\.codex\\skills")
    parser.add_argument("--message", help="Commit message for milestone publish")
    parser.add_argument("--check-only", action="store_true", help="Only initialize/check repo metadata")
    parser.add_argument("--skip-push", action="store_true", help="Commit only, do not push")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    result: dict[str, object] = {
        "repo_root": str(REPO_ROOT),
        "initialized": False,
        "remote_name": REMOTE_NAME,
        "remote_url": REMOTE_URL,
        "check_only": args.check_only,
    }

    try:
        initialized = ensure_repo_initialized()
        result["initialized"] = initialized
        ensure_local_config()
        result["user_name"] = GIT_USER_NAME
        result["user_email"] = GIT_USER_EMAIL
        result.update(ensure_remote())
        result["branch"] = current_branch() or "main"
        result["status_lines"] = working_tree_changes()

        if args.check_only:
            print(json.dumps(result, ensure_ascii=False))
            return 0

        if not args.message:
            raise GitPublisherError("缺少 --message，無法建立可追蹤的里程碑 commit")

        result.update(commit_and_push(args.message, args.skip_push))
        print(json.dumps(result, ensure_ascii=False))
        return 0
    except GitPublisherError as exc:
        result["error"] = str(exc)
        print(json.dumps(result, ensure_ascii=False))
        return 1
    except subprocess.CalledProcessError as exc:
        result["error"] = exc.stderr.strip() or exc.stdout.strip() or str(exc)
        print(json.dumps(result, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main())
