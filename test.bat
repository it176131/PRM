echo off
call conda activate prm
call python test.py
call conda deactivate