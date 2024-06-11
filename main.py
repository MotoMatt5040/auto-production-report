import looprunner
import traceback

if __name__ == '__main__':
    try:
        looprunner.run_loop()
    except Exception as err:
        print(traceback.format_exc())
        input("Press enter to continue")

    input("Press enter to continue")