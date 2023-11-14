import sys
import time
import helper
import constants
import pandas as pd
from selenium.webdriver.common.by import By

class main:
    hc = helper.HelperClass()
    counter = 2 # helps us keep track of what excel row we should be on

    # Take directory
    while True:
        try:
            dir = sys.argv[1]
            break
        except:
            print("[ERROR] Please pass in directory name.")
            print("""[INFO] Example: python main.py 'C:\\Users\\user1\\Documents\\file.xlsx'\n""")
            exit()

    # Open the file for reading & store column data
    try:
        file_name = dir
        hc.setFilePath(file_name)
        projectNamesList = pd.read_excel(io=file_name, header=None, skiprows=1, usecols="A").values.tolist()
        gogsLinksList = pd.read_excel(io=file_name, header=None, skiprows=1, usecols="B").values.tolist()
        print("\n>>>>> Found data file, starting migration. <<<<<\n")
    except:
        print("[ERROR] Failed to open file, please validate file existence/path.")
        exit()

    # Initialize chrome driver & login to GitLab
    driver = hc.initializeChromeDriver(constants.GITLAB_URL)
    hc.loginGitLab(driver)

    # Enter the name of the project in the search bar to verify it doesn't already exists, if it does, skip to next project
    for gogsLink, projName in zip(gogsLinksList, projectNamesList):

        # Search for the project name via search bar to check if it exists already
        projectSearchBar = driver.find_element(By.XPATH, "//input[@data-qa-selector='groups_filter_field']")
        projectSearchBar.send_keys(projName)
        time.sleep(2)

        # Verify project doesn't already exists
        if (constants.NO_RESULTS_FOUND in driver.page_source):
            hc.createProject(driver, counter, gogsLink, projName)
        else:
            # Make a list of all the reported projects
            projects_reported = driver.find_elements(By.XPATH, "//a[@data-testid='group-name']")
            projectExists = False
        
            # Look through projects if we get a match
            for project in projects_reported:
                if (str(projName)[2:-2] == project.text):
                    projectExists = True
                    break

            # If project does exists, skip to next project, else create it
            if (projectExists == True):
                # Let user know searched project already exists, continue to next project
                hc.writeToExcel(counter, projName, True)
                print(">>>>> Project", projName, "already exists, skipping to next project.\n")
                projectSearchBar.clear()
                time.sleep(5)
            else:
                hc.createProject(driver, counter, gogsLink, projName)
        counter += 1
        
    # Now update the configurations for all projects on Jenkins
    hc.modifyJenkinsProject(projectNamesList)

    # End
    driver.close()