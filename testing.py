"""
testing.py

This is the main driver used for testing ThirdEyeNav module
"""

from src.third_eye_nav import ThirdEyeNav
import time
import sys

def main():
    
    te = ThirdEyeNav('head', '', '')
    te.login()
    
    info = [True] * 12

    whatisthis = te.get_info('111111', info)
    print("This is what it is!")
    print(whatisthis)
    time.sleep(3)
    

if __name__ == '__main__':
    main()
    