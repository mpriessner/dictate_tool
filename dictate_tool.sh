#!/bin/bash

# Initialize conda
eval "$(/opt/homebrew/Caskroom/miniconda/base/bin/conda shell.bash hook)"

# Script to activate conda environment and run dictation tool
conda activate dict_mac_env
python /Users/mpriessner/windsurf_repos/dictate_tool/voice_shortcuts.py