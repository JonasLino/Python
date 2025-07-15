import time
import pyautogui

def glpi():
    positions = [(52, 761), (216, 756), (447, 758)]
    counter = 0
    
    while True:
        for pos in positions:
            pyautogui.click(x=pos[0], y=pos[1])
            time.sleep(60)
            counter += 1
            
            if counter >= 10:
                pyautogui.hotkey("ctrl", "R")
                counter = 0

def main():
    glpi()

if __name__ == "__main__":
    main()
