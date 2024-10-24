import subprocess
import pygetwindow as gw
import time
import psutil
import pyautogui
import win32gui
import win32con
import win32api
from PIL import Image




def open_bluestacks():
    # Path to the Bluestacks executable, adjust this to your specific installation path
    bluestacks_path = r"C:\Program Files\BlueStacks_nxt\HD-MultiInstanceManager.exe"

    try:
        # Open Bluestacks by executing the file
        subprocess.Popen([bluestacks_path])
        print("Bluestacks opened successfully.")
    except FileNotFoundError:
        print("Bluestacks executable not found. Please check the path.")
    except Exception as e:
        print(f"An error occurred: {e}")


def get_window_handle_by_title(window_title):
    # Find the window handle using its title
    handle = win32gui.FindWindow(None, window_title)
    if handle == 0:
        print(f"Window with title '{window_title}' not found.")
        return None
    return handle


def bring_window_to_foreground(handle):
    # Bring the window to the foreground using Windows API functions
    try:
        # If the window is minimized, restore it
        if win32gui.IsIconic(handle):
            win32gui.ShowWindow(handle, win32con.SW_RESTORE)

        # Bring the window to the foreground and set it as the active window
        win32gui.SetForegroundWindow(handle)

        # Additional step to ensure the window gets focus (sometimes needed)
        win32gui.ShowWindow(handle, win32con.SW_SHOW)
        time.sleep(0.5)

        print("Window brought to foreground successfully.")
    except Exception as e:
        print(f"Failed to bring window to foreground: {e}")


def activate_multi_instance_manager(
        window_title="BlueStacks Multi Instance Manager"):  # This will bring multi instance manager to the foreground
    time.sleep(5)

    handle = get_window_handle_by_title(window_title)

    if handle:
        bring_window_to_foreground(handle)
    else:
        print("Could not find the Multi Instance Manager window.")


def open_account(x, y):
    # first find windows coordinates
    window_title = 'BlueStacks Multi Instance Manager'

    try:
        window = gw.getWindowsWithTitle(window_title)[0]  # Get the first matching window
        # print(f"Window found: {window.title}")
        # print(f"Coordinates: {window.left}, {window.top}, {window.width}, {window.height}")
        # print(f"Clicking at: {absolute_x}, {absolute_y_4}")

        # Optional: Wait a moment before performing the click
        time.sleep(1)

        # Move the mouse to the calculated position and click
        pyautogui.moveTo(window.left + x, window.top + y, duration=1)
        pyautogui.click()


    except IndexError:
        print("Window not found.")


def zoom_in():
    with pyautogui.hold('ctrl'):
        pyautogui.moveTo(1000, 500, duration=0.5)
        pyautogui.scroll(100)

# open_account()
def open_all_accs():
    y_values = [110, 160, 200, 245, 290]
    for i in y_values:
        activate_multi_instance_manager()
        time.sleep(2)
        open_account(430, i)
        time.sleep(5)


# open_windows = gw.getAllTitles()
# print("Open windows:")
# for title in open_windows:
#     print(title)

def click_image(image_path, confidence=0.5):
    # Locate the image on the screen
    location = pyautogui.locateOnScreen(image_path, confidence=confidence, grayscale=True)
    if location:
        pyautogui.moveTo(location.left + location.width / 2, location.top + location.height / 2)
        pyautogui.click()
        return True
    else:
        print(f"Could not find {image_path}")
        return False


def open_game(number):
    activate_multi_instance_manager(f"Michail {number}")
    window = gw.getWindowsWithTitle(f"Michail {number}")[0]
    time.sleep(2)

    click_image("game_icon.png")


start_time = time.time()

# open_bluestacks()
# open_all_accs()
# time.sleep(15)
# for i in range(1,6):
#     open_game(i)
#     time.sleep(10)

end_time = time.time()
elapsed_time = end_time - start_time

def maximize_and_restore_window():
    pyautogui.press('f11')

def exploration_reset(num,sleep_time=3):
    exp_list = [
        (700, 1020),
        (1160, 700),
        (950, 770),
        (950, 40),
        (690, 30)
    ]

    activate_multi_instance_manager(f"Michail {num}")

    maximize_and_restore_window()
    time.sleep(sleep_time)

    for i in range(5):
        current = exp_list[i]
        pyautogui.moveTo(current[0], current[1], duration=0.5)
        pyautogui.click()
        time.sleep(sleep_time)

    maximize_and_restore_window()

boss_blacklist = []
missions = []
#
# for i in range(1,6): #Works
#     exploration_reset(i,5)

def intel(num, sleep_time=3):
    activate_multi_instance_manager(f"Michail {num}")
    maximize_and_restore_window()
    time.sleep(sleep_time)



    def world_check():
        pix = pyautogui.pixel(760,24)
        return pix

    # if not world_check() == (202, 232, 255):
    #     pyautogui.moveTo(1200,1030, duration=0.5)
    #     pyautogui.click()
    #     time.sleep(sleep_time)

    def intel_menu():
        pyautogui.moveTo(1210,730, duration=0.5)
        pyautogui.click()
        time.sleep(sleep_time)

    # pyautogui.moveTo(1210,730, duration=0.5)
    # intel_menu()
    # time.sleep(1)

    # Define RGB color for the boss (gold color)
    boss_color = (255,175,45)
    regular_colors = [(255, 255, 255), (237, 242, 245)]

    # Define an area of the screen to scan (adjust these values to your actual screen)
    scan_area = (650, 110, 1260, 1035)



    def solve_mission(x,y):
        pyautogui.moveTo(960,780, duration=0.5) #click on view
        pyautogui.click()

        if not pyautogui.pixelMatchesColor(972, 365, (170, 96, 6)):  # it is NOT a boss
            if pyautogui.pixelMatchesColor(734,1033, (255, 255, 255)): #20 is visible in intel
                print("Not a mission")
                pass
            time.sleep(2)
            pyautogui.moveTo(960, 620, duration=0.5)  # click on action
            pyautogui.click()
            time.sleep(3)

            if pyautogui.pixelMatchesColor(1230,455, (13, 50, 95)): #it's a monster

                time.sleep(2)
                pyautogui.moveTo(855,215, duration=0.5) #remove all heroes
                pyautogui.click()
                time.sleep(1)
                pyautogui.moveTo(800, 320, duration=0.5)  # start choosing hero
                pyautogui.click()
                time.sleep(1)
                heroes_ss = pyautogui.screenshot(region=(710, 480, 1210, 780)) #heroes array
                gina = pyautogui.locateOnScreen('gina.png', region=(710, 480, 1210, 780), grayscale=True, confidence=0.8)
                if gina:
                    pyautogui.moveTo(gina, duration=0.5)
                    pyautogui.click() #choose gina
                    pyautogui.moveTo(1090, 850, duration=0.5)  # triple assign
                    pyautogui.click()
                    time.sleep(1)
                    pyautogui.click()
                    time.sleep(1)
                    pyautogui.click()
                    time.sleep(1)
                    pyautogui.moveTo(970, 30, duration=0.5) # remove from choosing heroes
                    pyautogui.click()
                    time.sleep(2)

                    pyautogui.moveTo(1120, 1020, duration=0.5)  # click on deploy
                    pyautogui.click()

                    time.sleep(5)
                    while pyautogui.pixelMatchesColor(684, 223, (255, 255, 255)): #wait for march to return
                        time.sleep(1)
                    print("completed monster")
                    missions.append("monster")


                else:
                    print("no gina")

            elif pyautogui.pixelMatchesColor(749,120, (230, 242, 255)): #it's arenas

                time.sleep(2)
                pyautogui.moveTo(1160, 1010, duration=0.5)  # click fight
                pyautogui.click()

                while not pyautogui.pixelMatchesColor(940, 820, (59, 34, 28)): #wait for fight to finish
                    time.sleep(1)

                time.sleep(2)
                pyautogui.moveTo(950, 1000, duration=0.5)  # remove from reward
                pyautogui.click()
                print("completed arenas")
                missions.append("arenas")

                time.sleep(1)

            elif pyautogui.pixelMatchesColor(841,225, (255, 255, 255)): #it's tent
                pyautogui.moveTo(960, 530, duration=0.5)  # click on action
                time.sleep(2)
                while pyautogui.pixelMatchesColor(848, 226, (255, 255, 255)):  # wait for march to return
                    time.sleep(1)
                print("completed tent")
                missions.append("tent")
            else:
                print("no mission")
                pass

        else:
            boss_blacklist.append((x, y))
            print("Boss found")
            boss = True
            pass

    def find_mission():
        done = False
        boss = False
        img = pyautogui.screenshot()
        scan_area = (650, 150, 1260, 900)
        for x in range(670, 1225,5):
            if done:
                break
            for y in range(190, 910,5):
                if img.getpixel((x, y)) == (255, 255, 255) and (x, y) not in boss_blacklist: #if white
                    print(x, y)
                    pyautogui.moveTo(x, y, duration=0.5)
                    pyautogui.click()
                    time.sleep(1)
                    img = pyautogui.screenshot()
                    if img.getpixel((734, 1033)) != (255, 255, 255): #if not white
                        time.sleep(1)
                        print("before solving")
                        solve_mission(x,y)
                        if not boss:
                            print("after solving")
                            time.sleep(2)
                            intel_menu()
                            time.sleep(1)
                            pyautogui.moveTo(x, y, duration=0.5)
                            time.sleep(1)
                            pyautogui.click()
                        time.sleep(1)
                        pyautogui.moveTo(970, 30, duration=0.5)  # remove from choosing heroes
                        pyautogui.click()
                        time.sleep(1)
                        pyautogui.click(690,30)
                        time.sleep(1)
                        done = True
                        boss = False
                        break
                    else:
                        print("still in intel")



    # for i in range(5):
    #     zoom_in()
    #     time.sleep(0.5)
    time.sleep(1)
    intel_menu()
    time.sleep(2)
    find_mission()


    maximize_and_restore_window()
    print(boss_blacklist)


def alliance_contributions(num):
    activate_multi_instance_manager(f"Michail {num}")
    maximize_and_restore_window()
    time.sleep(1)
    pyautogui.moveTo(1100,1025, duration=0.5)
    pyautogui.click()
    time.sleep(3)
    pyautogui.moveTo(1100,800, duration=0.5)
    pyautogui.click()
    time.sleep(3)
    pyautogui.moveTo(950,880, duration=0.5)
    pyautogui.click()
    time.sleep(3)
    pyautogui.moveTo(1100,870, duration=0.5)
    pyautogui.click(clicks=50, interval=0.1)
    time.sleep(3)
    pyautogui.moveTo(690,30, duration=0.5)
    pyautogui.click(clicks=3, interval=0.5)
    time.sleep(1)
    maximize_and_restore_window()



def vip(num):
    activate_multi_instance_manager(f"Michail {num}")
    maximize_and_restore_window()
    time.sleep(6)
    pyautogui.moveTo(1080,60, duration=0.5)
    pyautogui.click()
    time.sleep(3)
    pyautogui.moveTo(1150,700, duration=0.5)
    pyautogui.click(clicks=2, interval=2)
    time.sleep(3)
    pyautogui.moveTo(1180,240, duration=0.5)
    pyautogui.click(clicks=2, interval=2)
    time.sleep(3)
    pyautogui.moveTo(690,30, duration=0.5)
    pyautogui.click()
    time.sleep(1)
    maximize_and_restore_window()

def recruit(num):
    activate_multi_instance_manager(f"Michail {num}")
    maximize_and_restore_window()
    time.sleep(2)
    pyautogui.moveTo(800,1030, duration=0.5)
    pyautogui.click()
    time.sleep(3)
    pyautogui.moveTo(1100,1020, duration=0.5)
    pyautogui.click()
    time.sleep(3)
    pyautogui.moveTo(820,710, duration=0.5)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(950, 20, duration=0.5)
    while not pyautogui.pixelMatchesColor(840,30,(89,122,147)):
        pyautogui.click()
        time.sleep(0.1)
    time.sleep(1)
    pyautogui.moveTo(690,30, duration=0.5)
    pyautogui.click()
    time.sleep(2)
    pyautogui.moveTo(690, 30, duration=0.5)
    pyautogui.click()
    time.sleep(1)
    maximize_and_restore_window()


def auto_join(num, state):
    activate_multi_instance_manager(f"Michail {num}")
    maximize_and_restore_window()
    time.sleep(1)
    pyautogui.moveTo(1100, 1025, duration=0.5)
    pyautogui.click()
    time.sleep(3)
    pyautogui.moveTo(820, 560, duration=0.5)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(950, 1030, duration=0.5)
    pyautogui.click()
    time.sleep(1)
    if state == "on":
        pyautogui.moveTo(1090, 920, duration=0.5)
        pyautogui.click()
    elif state == "off":
        pyautogui.moveTo(800, 920, duration=0.5)
        pyautogui.click()
    else:
        print("invalid state")
    time.sleep(1)
    pyautogui.moveTo(690, 30, duration=0.5)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(690, 30, duration=0.5)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(690, 30, duration=0.5)
    pyautogui.click()

    maximize_and_restore_window()


while True:
    intel(2)