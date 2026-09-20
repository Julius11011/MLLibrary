#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MLC Audio Podcast GitHub Release Manager & Uploader
===================================================
Automatically manages and syncs large MP3 study podcasts to GitHub Releases,
enabling free, high-speed CDN audio streaming on Cloudflare and mobile devices.

Usage:
  python upload_audio_to_github_release.py [--token YOUR_GITHUB_TOKEN] [--tag audio-v1]
"""

import os
import sys
import argparse
import webbrowser
from pathlib import Path
import urllib.request
import urllib.parse
import json

REPO_OWNER = "Julius11011"
REPO_NAME = "MLLibrary"
DEFAULT_TAG = "audio-v1"

def find_all_subject_mp3s(root_dir):
    subjects_dir = Path(root_dir) / "First Sem 1st Year" / "Subjects"
    if not subjects_dir.exists():
        subjects_dir = Path(root_dir) / "Subjects"
    
    mp3_files = []
    if subjects_dir.exists():
        for p in subjects_dir.rglob("*.mp3"):
            if not p.name.startswith(("~$", ".~")):
                mp3_files.append(p)
    return sorted(mp3_files, key=lambda x: x.name)

def get_or_create_release(token, tag=DEFAULT_TAG):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "User-Agent": "MLC-Audio-Sync"
    }
    
    # Check if release exists
    get_url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/tags/{tag}"
    req = urllib.request.Request(get_url, headers=headers)
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode())
            return data
    except urllib.error.HTTPError as e:
        if e.code != 404:
            print(f"[ERROR] Failed to fetch release: {e}")
            return None

    # Create release if not found
    post_url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases"
    payload = {
        "tag_name": tag,
        "name": f"MLC Audio Podcasts ({tag})",
        "body": "Studio-quality law school audio podcasts and spoken legal case digest compendiums for MLC Juris Doctor students.",
        "draft": False,
        "prerelease": False
    }
    req = urllib.request.Request(post_url, data=json.dumps(payload).encode(), headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode())
            print(f"[SUCCESS] Created GitHub Release: {tag}")
            return data
    except Exception as e:
        print(f"[ERROR] Failed to create release: {e}")
        return None

def upload_mp3_to_release(token, upload_url_template, mp3_path):
    upload_url = upload_url_template.split('{')[0]
    filename = mp3_path.name
    query = urllib.parse.urlencode({"name": filename})
    target_url = f"{upload_url}?{query}"
    
    size_mb = mp3_path.stat().st_size / (1024 * 1024)
    print(f"  -> Uploading: {filename} ({size_mb:.1f} MB)...")
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "audio/mpeg",
        "User-Agent": "MLC-Audio-Sync"
    }
    
    with open(mp3_path, "rb") as f:
        file_bytes = f.read()
        
    req = urllib.request.Request(target_url, data=file_bytes, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode())
            download_url = data.get("browser_download_url")
            print(f"  [DONE] Live CDN URL: {download_url}")
            return True
    except urllib.error.HTTPError as e:
        if e.code == 422:
            print(f"  [INFO] {filename} is already attached to this release.")
            return True
        else:
            print(f"  [ERROR] Upload failed ({e.code}): {e.read().decode()}")
            return False
    except Exception as e:
        print(f"  [ERROR] Upload error: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description="Upload MLC Audio Podcasts to GitHub Releases")
    parser.add_argument("--token", default=os.environ.get("GITHUB_TOKEN"), help="GitHub Personal Access Token (PAT)")
    parser.add_argument("--tag", default=DEFAULT_TAG, help="Release tag name (default: audio-v1)")
    parser.add_argument("--open-browser", action="store_true", help="Open GitHub Release page in default web browser")
    args = parser.parse_args()

    root_dir = Path(__file__).parent.resolve()
    mp3s = find_all_subject_mp3s(root_dir)

    print("===============================================================================")
    print("           MLC AUDIO PODCAST GITHUB RELEASE CLOUD SYNC")
    print("===============================================================================\n")

    if not mp3s:
        print("[INFO] No .mp3 files found in Subjects folder.")
        return

    print(f"Found {len(mp3s)} audio podcasts in Subjects:\n")
    for i, p in enumerate(mp3s, 1):
        mb = p.stat().st_size / (1024 * 1024)
        print(f"  {i}. {p.name} ({mb:.1f} MB)")
        print(f"     Path: {p}")

    print("\n-------------------------------------------------------------------------------")
    
    release_url = f"https://github.com/{REPO_OWNER}/{REPO_NAME}/releases/new"

    if args.token:
        print(f"\n[AUTH] Authenticating with GitHub API for tag '{args.tag}'...")
        release_data = get_or_create_release(args.token, args.tag)
        if release_data:
            upload_url_template = release_data.get("upload_url")
            existing_assets = {a["name"]: a["browser_download_url"] for a in release_data.get("assets", [])}
            
            print(f"\n[SYNC] Syncing {len(mp3s)} audio files to GitHub Release...")
            success_count = 0
            for mp3_file in mp3s:
                if mp3_file.name in existing_assets:
                    print(f"  [ALREADY SYNCED] {mp3_file.name}")
                    success_count += 1
                else:
                    if upload_mp3_to_release(args.token, upload_url_template, mp3_file):
                        success_count += 1
            print(f"\n[COMPLETED] Successfully synced {success_count}/{len(mp3s)} audio podcasts to cloud CDN!")
        else:
            print("[WARN] Could not access release via API. Opening browser release page instead...")
            webbrowser.open(release_url)
    else:
        print("\n[INFO] To upload automatically via script, set GITHUB_TOKEN or pass --token YOUR_TOKEN.")
        print(f"[MANUAL UPLOAD] Opening GitHub Release page: {release_url}")
        print("\nSteps:")
        print(f"  1. Set Tag: {args.tag}")
        print("  2. Drag and drop the MP3 files listed above")
        print("  3. Click 'Publish release'")
        print("-------------------------------------------------------------------------------")
        try:
            webbrowser.open(release_url)
        except Exception:
            pass

if __name__ == "__main__":
    main()
