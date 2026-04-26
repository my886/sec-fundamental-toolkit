"""Smoke test: are prints even reaching the terminal?"""
import sys
print("Step 1: print() works", flush=True)
sys.stdout.write("Step 2: sys.stdout.write works\n")
sys.stdout.flush()

print("Step 3: importing dotenv...", flush=True)
from dotenv import load_dotenv
load_dotenv()

print("Step 4: importing pandas...", flush=True)
import pandas as pd

print("Step 5: importing edgar...", flush=True)
from edgar import Company

print("Step 6: all imports done. If you see this, the script and Python are fine.", flush=True)
