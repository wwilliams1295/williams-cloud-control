# tools/auto_improver.py
from __future__ import annotations

import argparse
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
from typing import Dict, List, Optional, Tuple

# -------------------------------------------------------------------
# Project root and .env loading (ensures API keys are visible)
# -------------------------------------------------------------------
ROOT = pathlib.Path(__file__).resolve().parents[1]

try:
    from dotenv import load_dotenv  # type: ignore
except Exception:
    load_dotenv = None

if load_dotenv is not None:
    # Force-load the repo .env so child modules (e.g., ai_patch) see keys at import-time
    load_dotenv(dotenv_path=str(ROOT / ".env"), override=True)

# Import after .env load so any module-level reads see the variables.
from ai_patch import ask_for_patch  # noqa: E402  # imported after .env load on purpose

try:
    import yaml  # type: ignore
except Exception:
    yaml = None  # Policy will fall back to defaults

# -------------------------------------------------------------------
# Policy
# -------------------------------------------------------------------
POLICY_PATH = ROOT / "tools" / "policy.yaml"
DEFAULT_POLICY: Dict[str, object] = {
    "max_changed_lines": 400,
    "allow_new_files": True,
    "allowed_paths": [
        "plugins/**",
        "tests/**",
        "tools/**",
        "core/**",
        "functions.py",
        "mailer.py",
        "cloud.py",
        "agent.py",
    ],
    "forbidden_patterns": ["shell=True"],
}
if yaml and POLICY_PATH.exists():
    try:
        POLICY: Dict[str, object] = (
            yaml.safe_load(POLICY_PATH.read_text(encoding="utf-8")) or DEFAULT_POLICY
        )
    except Exception:
        POLICY = DEFAULT_POLICY
else:
    POLICY = DEFAULT_POLICY

# Only lint/test our own code — never venv/site-packages/backups/etc.
PROJECT_TARGETS: List[str] = [
    "agent.py",
    "cloud.py",
    "functions.py",
    "mailer.py",
    "core",
    "plugins",
    "tools",
    "tests",
]


# -------------------------------------------------------------------
# Small helpers
# -------------------------------------------------------------------
def run(cmd: List[str]) -> int:
    """Run a command in the repo root, echoing it (no custom cwd)."""
    print("+", " ".join(cmd))
    return subprocess.call(cmd, cwd=str(ROOT))


def ensure_bandit_config() -> None:
    """Create a minimal bandit.yaml if missing, so bandit doesn't traverse deps."""
    path = ROOT / "bandit.yaml"
    if not path.exists():
        path.write_text(
            "exclude_dirs:\n"
            "  - venv\n"
            "  - .venv\n"
            "  - backups\n"
            "  - node_modules\n"
            '  - "**/site-packages"\n'
            '  - "**/.tox"\n'
            '  - "**/.mypy_cache"\n'
            '  - "**/.pytest_cache"\n'
            "severity: LOW\n"
            "confidence: HIGH\n",
            encoding="utf-8",
        )


def checks_pass(tmp_cwd: Optional[pathlib.Path] = None) -> bool:
    """
    Run quick static checks on our code only (exclude venv/backups/etc).
    Uses config files if present (.ruff.toml, mypy.ini, bandit.yaml).
    """
    ensure_bandit_config()
    cwd = str(tmp_cwd or ROOT)
    print(f"[auto_improver] running checks in {cwd}")
    rc = 0

    # Ruff: lint + fix + format just our targets
    rc |= subprocess.call(
        [sys.executable, "-m", "ruff", "check", "--fix", *PROJECT_TARGETS], cwd=cwd
    )
    rc |= subprocess.call(
        [sys.executable, "-m", "ruff", "format", *PROJECT_TARGETS], cwd=cwd
    )

    # Mypy: relies on mypy.ini; keep lenient; only our targets
    rc |= subprocess.call(["mypy", *PROJECT_TARGETS], cwd=cwd)

    # Bandit: respects bandit.yaml excludes; only our targets
    rc |= subprocess.call(["bandit", "-q", "-r", *PROJECT_TARGETS], cwd=cwd)

    # Tests (optional)
    if (pathlib.Path(cwd) / "tests").exists():
        rc |= subprocess.call(["pytest", "-q"], cwd=cwd)

    return rc == 0


# -------------------------------------------------------------------
# Patch normalization & validation utilities
# -------------------------------------------------------------------
_FENCE_RE = re.compile(r"^```(?:diff)?\s*$|^```\s*$", re.M)
_UNICODE_SMARTS: Dict[str, str] = {
    "\u201c": '"',
    "\u201d": '"',
    "\u2018": "'",
    "\u2019": "'",
    "\u2013": "-",
    "\u2014": "-",
    "\u00a0": " ",
}


def _sanitize_text(raw: str) -> str:
    """Strip code fences, normalize Unicode punctuation, force LF, ensure trailing newline."""
    if not raw:
        return ""
    s = _FENCE_RE.sub("", raw)
    for k, v in _UNICODE_SMARTS.items():
        s = s.replace(k, v)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    if not s.endswith("\n"):
        s += "\n"
    return s


def normalize_patch(raw: str) -> str:
    """
    Strip Markdown fences and non-diff noise, keep only valid `diff --git` blocks.
    """
    s = _sanitize_text(raw).strip()

    # Extract *only* unified-diff file blocks
    blocks = re.findall(
        r"(?:^|\n)(diff --git .+?)(?=\n(?:diff --git |\Z))", s, re.DOTALL
    )
    if not blocks:
        # If the string already starts with diff --git and has no multiple blocks
        if s.startswith("diff --git "):
            return s
        return ""  # nothing usable

    cleaned = "\n".join(
        b.strip() for b in blocks if b.strip().startswith("diff --git ")
    )
    return cleaned.strip()


def patch_allowed_verbose(patch: str, policy: dict) -> Tuple[bool, str]:
    """
    Validate a proposed patch against policy, returning (ok, reason).

    Key behaviors:
    - Only checks forbidden_patterns against ADDED lines (lines starting with '+' and not '+++').
    - Validates paths via file headers (---/+++).
    """
    if not patch or not patch.strip().startswith("diff --git"):
        return False, "Not a unified diff (missing 'diff --git')."

    max_lines = int(policy.get("max_changed_lines", 400))
    total_lines = patch.count("\n")
    if total_lines > max_lines + 50:
        return False, f"Patch too large: {total_lines} lines > {max_lines + 50} limit."

    allow_new = bool(policy.get("allow_new_files", False))
    allowed_paths = policy.get("allowed_paths", []) or []
    forbidden = policy.get("forbidden_patterns", []) or []

    def ok_path(path: str) -> bool:
        for pat in allowed_paths:
            try:
                if pathlib.PurePosixPath(path).match(pat):
                    return True
            except Exception:
                continue
        return False

    # Validate file paths from headers and new-file rule
    for line in patch.splitlines():
        if line.startswith(("--- ", "+++ ")):
            if line.strip().endswith("/dev/null"):
                continue
            # header looks like: '--- a/path' or '+++ b/path'
            hdr_path = line.split("\t")[-1].split(" ", 1)[-1]
            hdr_path = hdr_path.replace("a/", "").replace("b/", "").strip()
            if not ok_path(hdr_path):
                return False, f"Path not allowed by policy: {hdr_path}"
        if line.startswith("new file mode") and not allow_new:
            return False, "New files are not allowed by policy."

    # Check forbidden patterns ONLY on lines being added
    added_lines: List[str] = []
    for line in patch.splitlines():
        if line.startswith("+") and not line.startswith("+++"):
            added_lines.append(line[1:])  # strip leading '+'

    added_blob = "\n".join(added_lines)
    for bad in forbidden:
        if bad and bad in added_blob:
            return False, f"Forbidden pattern found in added code: {bad!r}"

    return True, "OK"


# -------------------------------------------------------------------
# Patch preflight and application (with robust fallbacks)
# -------------------------------------------------------------------
def check_patch_with_git(patch_text: str) -> Tuple[bool, str]:
    """
    Run `git apply --check` to validate the patch without applying it.
    Returns (ok, stderr_tail).
    """
    with tempfile.NamedTemporaryFile("w", delete=False, encoding="utf-8") as f:
        f.write(patch_text)
        temp_path = f.name
    try:
        proc = subprocess.run(
            ["git", "apply", "--check", temp_path],
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            check=False,
        )
        ok = proc.returncode == 0
        stderr_tail = (proc.stderr or "")[-2000:]
        return ok, stderr_tail
    finally:
        try:
            os.unlink(temp_path)
        except OSError:
            pass


def _coerce_file_blocks_to_diff(model_text: str) -> str:
    """
    Convert blocks of the form:
        FILE: relative/path.py
        <entire new file contents...>
        --- endfile
    into a proper unified diff using `git diff --no-index`.
    """
    s = _sanitize_text(model_text)
    lines = s.splitlines()
    out: List[Tuple[pathlib.Path, str]] = []
    i = 0
    while i < len(lines):
        if lines[i].startswith("FILE: "):
            rel = lines[i][6:].strip()
            i += 1
            buf: List[str] = []
            while i < len(lines) and lines[i].strip() != "--- endfile":
                buf.append(lines[i])
                i += 1
            # consume any number of --- endfile lines
            while i < len(lines) and lines[i].strip() == "--- endfile":
                i += 1
            path = ROOT / rel
            content = "\n".join(buf)
            if content and not content.endswith("\n"):
                content += "\n"
            out.append((path, content))
        else:
            i += 1
    if not out:
        return ""

    diffs: List[str] = []
    tmp_dir = ROOT / ".ai_patch_tmp"
    tmp_dir.mkdir(exist_ok=True)
    for path, new_contents in out:
        tmp_new = tmp_dir / (path.name + ".new")
        tmp_new.parent.mkdir(parents=True, exist_ok=True)
        tmp_new.write_text(new_contents, encoding="utf-8")
        old = str(path)
        new = str(tmp_new)
        proc = subprocess.run(
            ["git", "--no-pager", "diff", "--no-index", "--", old, new],
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            check=False,
        )
        if proc.stdout.strip():
            diffs.append(proc.stdout)
    return "\n".join(diffs)


def apply_patch(patch: str) -> bool:
    """
    Apply a unified diff. Prefer `git apply` in a repo;
    otherwise use `patch` with stdin.
    """
    in_repo = run(["git", "rev-parse", "--is-inside-work-tree"]) == 0
    if in_repo:
        with tempfile.NamedTemporaryFile("w", delete=False, encoding="utf-8") as f:
            f.write(patch)
            temp_path = f.name
        try:
            return run(["git", "apply", "--reject", "--whitespace=fix", temp_path]) == 0
        finally:
            try:
                os.unlink(temp_path)
            except OSError:
                pass

    # Fallback: use `patch` reading from stdin (no shell)
    try:
        proc = subprocess.run(
            ["patch", "-p1"],
            cwd=str(ROOT),
            input=patch,
            text=True,
            capture_output=True,
            check=False,
        )
        if proc.returncode != 0:
            print("[apply_patch] patch stderr:\n", (proc.stderr or "")[-2000:])
            print("[apply_patch] patch stdout:\n", (proc.stdout or "")[-2000:])
        return proc.returncode == 0
    except FileNotFoundError:
        print("[apply_patch] 'patch' tool not found on this system.")
        return False


def split_diff_blocks(patch_text: str) -> List[str]:
    """Split a multi-file diff into a list of per-file diff blocks."""
    if not patch_text:
        return []
    s = normalize_patch(patch_text)
    return re.findall(r"(?:^|\n)(diff --git .+?)(?=\n(?:diff --git |\Z))", s, re.DOTALL)


def apply_patch_block(block: str) -> bool:
    """Apply a single file's diff block with git apply --reject."""
    with tempfile.NamedTemporaryFile("w", delete=False, encoding="utf-8") as f:
        f.write(block)
        temp = f.name
    try:
        return run(["git", "apply", "--reject", "--whitespace=fix", temp]) == 0
    finally:
        try:
            os.unlink(temp)
        except OSError:
            pass


# -------------------------------------------------------------------
# Model interaction helpers
# -------------------------------------------------------------------
def pick_provider(default_choice: Optional[str]) -> Optional[str]:
    """Pick an available provider (openai/perplexity) based on keys present."""
    have_openai = bool(os.getenv("OPENAI_API_KEY"))
    have_pplx = bool(os.getenv("PPLX_API_KEY"))

    if default_choice == "openai" and have_openai:
        return "openai"
    if default_choice == "perplexity" and have_pplx:
        return "perplexity"
    if have_openai:
        return "openai"
    if have_pplx:
        return "perplexity"
    return None


def build_instructions() -> str:
    """Instruction string sent to the model (explicit edit boundaries)."""
    return (
        "Goal: Make small, surgical improvements in allowed files.\n"
        "- Propose a SINGLE, minimal unified diff per run.\n"
        "- Prefer touching ONE file only; avoid large refactors.\n"
        "- Keep existing context exact so `git apply` succeeds.\n"
        "- Do NOT include Markdown fences or prose, only the diff.\n\n"
        "EDIT BOUNDARIES:\n"
        "- Only modify: plugins/**, tests/**, tools/**, core/**, functions.py, mailer.py, cloud.py, agent.py\n"
        "- New files allowed only under: plugins/** or tests/**\n"
        "Constraints: keep diffs small; avoid secrets; avoid shell=True; use subprocess with shell=False.\n"
        "Return ONLY a strict unified diff (git apply format)."
    )


def ask_model_for_patch(args_provider: Optional[str]) -> str:
    """Call the model and return the proposed patch (may be empty)."""
    provider = pick_provider(args_provider)
    if not provider:
        print(
            "No model API keys detected (OPENAI_API_KEY/PPLX_API_KEY). Skipping AI patch step."
        )
        return ""

    instructions = build_instructions()
    try:
        if run(["git", "rev-parse", "--is-inside-work-tree"]) == 0:
            diff_hint = subprocess.check_output(
                ["git", "--no-pager", "diff"],
                cwd=str(ROOT),
            ).decode("utf-8", "ignore")
        else:
            diff_hint = ""
    except Exception as e:
        print(f"[ask_model_for_patch] Could not get diff hint: {e}")
        diff_hint = ""

    print(f"Using provider: {provider}")
    return ask_for_patch(instructions, diff_hint, provider=provider) or ""


# -------------------------------------------------------------------
# Orchestration
# -------------------------------------------------------------------
def call_local_auto_fix() -> None:
    """
    Optional: call tools/auto_fix.py (custom fixer).
    Only runs if requested via --auto-fix flag.
    """
    fixer = ROOT / "tools" / "auto_fix.py"
    if fixer.exists():
        run([sys.executable, str(fixer)])


def _safe_apply_script() -> Optional[pathlib.Path]:
    """Return path to tools/safe-apply-patch.sh if present and executable."""
    p = ROOT / "tools" / "safe-apply-patch.sh"
    try:
        if p.exists() and os.access(str(p), os.X_OK):
            return p
    except Exception:
        pass
    return None


def self_heal(args: argparse.Namespace) -> bool:
    """
    Try to automatically fix issues:
      1) Optionally run local auto_fix (--auto-fix pre).
      2) Re-run checks.
      3) If still failing, ask model for diff and try to apply.
      4) Prefer safe-apply script; else preflight + per-file salvage.
      5) Run Ruff fix/format; optionally run local auto_fix (--auto-fix post).
      6) Re-run checks after changes.
    """
    if args.auto_fix in ("pre", "both"):
        call_local_auto_fix()
        if checks_pass():
            return True

    patch_raw = ask_model_for_patch(args.provider)
    patch = normalize_patch(patch_raw)
    if not patch:
        # Fall back to FILE: blocks → unified diff
        coerced = _coerce_file_blocks_to_diff(patch_raw)
        if coerced:
            patch = normalize_patch(coerced)

    (ROOT / "tools" / "out").mkdir(parents=True, exist_ok=True)
    patch_path = ROOT / "tools" / "out" / "patch.diff"
    patch_path.write_text(patch or "", encoding="utf-8")

    ok, reason = patch_allowed_verbose(patch, POLICY)
    if not ok:
        print("No safe patch proposed:", reason)
        preview = "\n".join((patch or patch_raw).splitlines()[:80])
        print("Patch preview:\n", preview)
        return False

    # Prefer the self-healing shell script if available
    safe_apply = _safe_apply_script()
    if safe_apply:
        print("[auto_improver] attempting to apply patch via safe-apply-patch.sh")
        rc = subprocess.call([str(safe_apply), str(patch_path)], cwd=str(ROOT))
        if rc != 0:
            print(f"[auto_improver] safe-apply-patch exited with code {rc}")
        # Regardless of rc, continue to post-steps (Ruff etc.) and final checks
    else:
        # Built-in Python fallback:
        # Preflight check for the whole patch…
        ok_check, err = check_patch_with_git(patch)
        if not ok_check:
            print("Patch failed preflight (--check). git says:\n", err)

            # Second-pass cleanup (normalize again)
            patch2 = normalize_patch(patch)
            ok_check2, err2 = check_patch_with_git(patch2)
            if not ok_check2:
                print("Patch still invalid after cleanup. git says:\n", err2)

                # Try per-file salvage
                blocks = split_diff_blocks(patch or patch_raw)
                if not blocks:
                    return False

                failures = 0
                for i, b in enumerate(blocks, 1):
                    okb, errb = check_patch_with_git(b)
                    if not okb:
                        print(f"[block {i}] preflight failed, skipping.\n{errb}")
                        failures += 1
                        continue
                    if not apply_patch_block(b):
                        print(f"[block {i}] failed to apply.")
                        failures += 1

                if failures == len(blocks):
                    print("All blocks failed; giving up.")
                    return False
            else:
                # cleaned version preflight succeeded; continue with cleaned patch
                patch = patch2

        # Try full apply (Python path)
        if not apply_patch(patch):
            blocks = split_diff_blocks(patch or patch_raw)
            if not blocks:
                print("Patch failed to apply and could not split into blocks.")
                return False

            failures = 0
            for i, b in enumerate(blocks, 1):
                okb, errb = check_patch_with_git(b)
                if not okb:
                    print(f"[block {i}] preflight failed, skipping.\n{errb}")
                    failures += 1
                    continue
                if not apply_patch_block(b):
                    print(f"[block {i}] failed to apply.")
                    failures += 1

            if failures == len(blocks):
                print("All blocks failed; giving up.")
                return False

    # Always run Ruff after changes; it will auto-fix many issues and keep the tree clean.
    subprocess.call(
        [sys.executable, "-m", "ruff", "check", "--fix", *PROJECT_TARGETS],
        cwd=str(ROOT),
    )
    subprocess.call(
        [sys.executable, "-m", "ruff", "format", *PROJECT_TARGETS], cwd=str(ROOT)
    )

    if args.auto_fix in ("post", "both"):
        call_local_auto_fix()

    return checks_pass()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument(
        "--provider",
        choices=["openai", "perplexity"],
        default=os.getenv("AI_PROVIDER", "openai"),
    )
    ap.add_argument(
        "--auto-fix",
        choices=["never", "pre", "post", "both"],
        default=os.getenv("AUTO_FIX_MODE", "never"),
        help="When to run tools/auto_fix.py (default: never)",
    )
    ap.add_argument("--show", action="store_true")
    ap.add_argument(
        "--no-docker",
        action="store_true",
        help="Force local sandbox instead of Docker, if desired.",
    )
    # Policy override knobs (optional)
    ap.add_argument(
        "--allow-new", action="store_true", help="Temporarily allow new files."
    )
    ap.add_argument("--max-lines", type=int, help="Temporarily raise patch size limit.")
    ap.add_argument(
        "--allow-path",
        action="append",
        default=[],
        help="Temporarily add allowed path glob (repeatable).",
    )
    args = ap.parse_args()

    # Apply temporary policy overrides from CLI
    if args.allow_new:
        POLICY["allow_new_files"] = True
    if args.max_lines:
        POLICY["max_changed_lines"] = int(args.max_lines)
    if args.allow_path:
        # POLICY["allowed_paths"] is List[str], ensure set merge operates on str
        POLICY["allowed_paths"] = list(
            set(map(str, POLICY.get("allowed_paths", [])))
            | set(map[str, args.allow_path])  # type: ignore[arg-type]
        )

    # Fast path: if baseline passes, great. Otherwise, try to self-heal.
    if not checks_pass():
        print("Baseline checks failed. Attempting self-heal...")
        if not self_heal(args):
            print("Self-heal failed. Aborting.")
            sys.exit(1)

    # Show last proposed patch (if any)
    if args.show:
        out = ROOT / "tools" / "out" / "patch.diff"
        if out.exists():
            print("\n===== BEGIN PROPOSED PATCH =====\n")
            print(out.read_text(encoding="utf-8"))
            print("\n=====  END  PROPOSED PATCH  =====\n")
        else:
            print("No patch file found at tools/out/patch.diff")

    # Dry run ends here unless --apply is set
    if args.dry_run and not args.apply:
        print("DRY RUN – baseline is green; no apply step requested.")
        sys.exit(0)

    # Sandbox run
    use_docker = False
    if not args.no_docker:
        try:
            subprocess.check_call(
                ["docker", "--version"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            use_docker = True
        except Exception:
            use_docker = False

    if use_docker:
        if (
            run(
                [
                    "docker",
                    "build",
                    "-f",
                    "Dockerfile.sandbox",
                    "-t",
                    "selfimprove:latest",
                    ".",
                ]
            )
            != 0
        ):
            print("Docker build failed.")
            sys.exit(1)

        if run(["docker", "run", "--rm", "selfimprove:latest"]) != 0:
            print("Sandbox tests failed, reverting if git.")
            if run(["git", "rev-parse", "--is-inside-work-tree"]) == 0:
                run(["git", "reset", "--hard", "HEAD"])
            sys.exit(1)

        print("All good; changes applied.")
        sys.exit(0)
    else:
        print("[auto_improver] running in local sandbox mode – no Docker")
        with tempfile.TemporaryDirectory() as td:
            tmp_repo = pathlib.Path(td) / "repo"
            # Copy the repo but ignore envs, caches, backups
            shutil.copytree(
                ROOT,
                tmp_repo,
                dirs_exist_ok=True,
                ignore=shutil.ignore_patterns(
                    "venv",
                    ".venv",
                    "backups",
                    ".git",
                    "__pycache__",
                    ".mypy_cache",
                    ".pytest_cache",
                    "node_modules",
                    "*site-packages*",
                ),
            )
            # Run checks *inside* the sandbox copy
            if not checks_pass(tmp_repo):
                print("[auto_improver] sandbox checks failed")
                sys.exit(1)

            # If sandbox is clean, our changes (already applied) are considered good.
            print("[auto_improver] sandbox check complete – applying to main tree")
            print("All good; changes applied.")
            sys.exit(0)


if __name__ == "__main__":
    main()
