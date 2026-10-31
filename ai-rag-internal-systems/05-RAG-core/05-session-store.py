#!/usr/bin/env python3
"""Per-session conversation memory backed by Redis with a TTL.

Tracker task: P5-11 (conversation memory)
Prompts: redis_host, redis_port, session_id, action (set/get/clear),
         message_text (for set), ttl_seconds
Prerequisites: Redis running (02-install-redis.sh), redis-cli on PATH.
  Sessions are short-lived by design; nothing is stored beyond the TTL
  unless the user explicitly saves the conversation elsewhere.
"""
import subprocess
import sys


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def prompt_int(label, default):
    while True:
        raw = input(f"{label} [{default}]: ").strip() or str(default)
        if raw.isdigit() and int(raw) > 0:
            return int(raw)
        print("Enter a positive whole number.")


def redis_cli(host, port, *args):
    result = subprocess.run(["redis-cli", "-h", host, "-p", str(port), *args],
                            capture_output=True, text=True)
    if result.returncode != 0:
        print(result.stderr)
        sys.exit(1)
    return result.stdout.strip()


def main():
    print("=== Step 1: Session settings ===")
    redis_host = prompt("Redis host", "PRD-AIRAG-DB01")
    redis_port = prompt_int("Redis port", 6379)
    session_id = prompt("Session ID")
    action = prompt("Action (set/get/clear)", "get").lower()
    ttl_seconds = prompt_int("Session TTL (seconds)", 3600)

    key = f"session:{session_id}"
    if action == "set":
        message_text = prompt("Message text to append")
        redis_cli(redis_host, redis_port, "RPUSH", key, message_text)
        redis_cli(redis_host, redis_port, "EXPIRE", key, str(ttl_seconds))
        print(f"Appended to {key} (TTL {ttl_seconds}s).")
    elif action == "get":
        messages = redis_cli(redis_host, redis_port, "LRANGE", key, "0", "-1")
        print(f"--- Session {session_id} ---")
        print(messages if messages else "(empty or expired)")
    elif action == "clear":
        redis_cli(redis_host, redis_port, "DEL", key)
        print(f"Session {session_id} cleared.")
    else:
        print("ERROR: action must be set, get or clear.")
        sys.exit(1)


if __name__ == "__main__":
    main()
