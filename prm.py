from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import os
from datetime import date
from time import sleep
import logging
import pandas as pd
from calendar import monthrange


def _get_date():
    """
    Get today's date, format, and return string

    Return
    ------
    today_str: str
    """

    logger = logging.getLogger("_get_date")

    logger.debug("get today's date")
    # get today's date
    today = date.today()

    logger.debug("convert to string and format as month/day/4-digit year")
    # convert to string and format as month/day/4-digit year
    today_str = today.strftime("%m/%d/%Y")

    logger.debug("return today_str")
    return today_str


def _get_end_of_month(today_str):
    """
    Get's the last day of the month, formats it, and returns as string

    Parameters
    ----------
    today_str: str
        Today's date string

    Returns
    -------
    end_of_month_str: str
    """

    logger = logging.getLogger("_get_end_of_month")

    logger.debug("convert today_str to datetime")
    # convert today_str to datetime
    today = pd.to_datetime(today_str, format="%m/%d/%Y")

    logger.debug("get end of month")
    # get end of month
    end_of_month = monthrange(today.year, today.month)[-1]

    logger.debug("create end of month string")
    # create end of month string
    end_of_month_str = f"{today.month}/{end_of_month}/{today.year}"

    logger.debug("return end_of_month_str")
    return end_of_month_str


def check_url(driver):
    """"""
    new_work_url = "https://ukhs.pvcloud.com/planview/ConfiguredScreens/ConfiguredScreen.aspx?sid=CfgDef$WDT&mode=RW&popup=1&back=close"

    return driver.current_url == new_work_url


def login(driver, wait):
    """
    Log-in to PRM

    Parameters
    ----------
    driver: selenium.webdriver.chrome.webdriver.WebDriver
        The driver used to control the browser

    wait: selenium.webdriver.support.wait.WebDriverWait
        Tells the driver to look for an element every 0.5 seconds until it is found

    Returns
    -------
    None
    """

    logger = logging.getLogger("login")

    logger.debug("define new work url")
    # define new work url
    new_work_url = "https://ukhs.pvcloud.com/planview/ConfiguredScreens/ConfiguredScreen.aspx?sid=CfgDef$WDT&mode=RW&popup=1&back=close"

    logger.debug("navigate to new work page")
    # navigate to new work page
    driver.get(new_work_url)

    sleep(2)

    if not check_url(driver):
        logger.debug("locate user name input")
        # locate user name input
        user_name_input = wait.until(
            EC.presence_of_element_located(
                # (By.CSS_SELECTOR, "input[name='UserName']")
                (By.CSS_SELECTOR, "input[type='email']")
            )
        )

        logger.debug("type user name")
        # type user name
        # user_name_input.send_keys(os.getenv("user"))
        user_name_input.send_keys(os.getenv("email"))

        logger.debug("click 'Next'")
        # click "Next"
        next_button = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[text()='Next']")
            )
        )
        next_button.click()

        logger.debug("locate password input")
        # locate password input
        password_input = wait.until(
            EC.presence_of_element_located(
                # (By.CSS_SELECTOR, "input[name='Password']")
                (By.CSS_SELECTOR, "input[type='password']")
            )
        )

        logger.debug("type password")
        # type password
        password_input.send_keys(os.getenv("pass"))

        logger.debug("locate submit button")
        # locate submit button
        submit_button = wait.until(
            EC.presence_of_element_located(
                # (By.CSS_SELECTOR, "span[class='submit'][role='button']")
                (By.CSS_SELECTOR, "button[type='submit']")
            )
        )

        logger.debug("click button")
        # click button
        submit_button.click()

    logger.debug("return None")
    return None


def new_work_page1(wait, today_str, description):
    """
    Fill in fields on first "New Work" page

    Parameters
    ----------
    wait: selenium.webdriver.support.wait.WebDriverWait
        Tells the driver to look for an element every 0.5 seconds until it is found

    today_str: str
        Today's date formatted as mm/dd/yyyy

    description: str
        Description/Name of project

    Returns
    -------
    None
    """

    logger = logging.getLogger("new_work_page1")

    logger.debug("locate workstream input")
    # locate workstream input
    workstream_input = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR,
             "div[class='attribute-part'][title='Description of the project parent in Work stucture']"
             "> span"
             "> input")
        )
    )

    logger.debug("type Business Intelligence")
    # type "Business Intelligence"
    workstream_input.send_keys("Business Intelligence")

    logger.debug("locate description input")
    # locate description input
    description_input = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[title='The description of the project, or the work entity.']")
        )
    )

    logger.debug(f"type {description}")
    # type <name of project>
    description_input.send_keys(f"{description}")

    logger.debug("locate work type selector")
    # locate work type selector
    work_type_selector = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "option[value='3036594']")
        )
    )

    logger.debug("click <option>")
    # click <option>
    work_type_selector.click()

    logger.debug("locate requested start input")
    # locate requested start input
    requested_start_input = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR,
             "input[title='The date the project has to be started. This date is used by the CPM forward pass for early date calculations. ']")
        )
    )

    sleep(1)

    logger.debug(f"type {today_str}")
    # type <date>
    requested_start_input.send_keys(today_str)

    logger.debug("locate requested finish input")
    # locate requested finish input
    requested_finish_input = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR,
             "input[title='The date the project has to be completed by. This date is used by the CPM backward pass for late date calculations. ']")
        )
    )

    logger.debug("get end_of_month_str")
    # get end_of_month_str
    end_of_month_str = _get_end_of_month(today_str)

    sleep(2)

    logger.debug(f"type {end_of_month_str}")
    # type <date>
    requested_finish_input.send_keys(end_of_month_str)

    sleep(1)

    logger.debug("locate save button")
    # locate save button
    save_button = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//*[text()='Save']")
        )
    )

    logger.debug("click button")
    # click button
    save_button.click()

    logger.debug("return None")
    return None


def new_work_page2(wait, today_str, bi_service_name):
    """
    Fill in fields on second "New Work" page

    Parameters
    ----------
    wait: selenium.webdriver.support.wait.WebDriverWait
        Tells the driver to look for an element every 0.5 seconds until it is found

    today_str: str
        Today's date formatted as mm/dd/yyyy

    bi_service_name: str
        Name of BI Service

    Returns
    -------
    None
    """

    logger = logging.getLogger("new_work_page2")

    logger.debug("locate bi service name selector")
    # locate bi service name selector
    bi_service_name_selector = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, f"//*[text()='{bi_service_name}']")
        )
    )

    logger.debug(f"click {bi_service_name}")
    # click <option>
    bi_service_name_selector.click()

    logger.debug("locate bi scoped date input")
    # locate bi scoped date input
    bi_scoped_date_input = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR,
             "div[class='attribute-field'][id='bi_scoped_date']"
             "> div"
             "> input")
        )
    )

    logger.debug(f"type {today_str}")
    # type <date>
    bi_scoped_date_input.send_keys(today_str)

    logger.debug("locate bi date approved input")
    # locate bi date approved input
    bi_date_approved_input = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR,
             "div[class='attribute-field'][id='bi_date_created']"
             "> div"
             "> input")
        )
    )

    logger.debug(f"type {today_str}")
    # type <date>
    bi_date_approved_input.send_keys(today_str)

    logger.debug("locate save and complete button")
    # locate save and complete button
    save_and_complete_button = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//*[text()='Save and Complete']")
        )
    )

    logger.debug("click button")
    # click button
    save_and_complete_button.click()

    logger.debug("locate action menu button")
    # locate action menu button
    action_menu_button = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR,
             "button[class='banner-title-bar-button']"
             "> span[title='Actions']")
        )
    )

    logger.debug("click button")
    # click button
    action_menu_button.click()

    logger.debug("locate work and assignments button")
    # locate work and assignments button
    work_and_assignments_button = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "li[title='Work and Assignments']")
        )
    )

    logger.debug("click button")
    # click button
    work_and_assignments_button.click()

    logger.debug("return None")
    return None


def _locate_grid_canvas_right(wait):
    """
    Locates the grid canvas right and returns it

    Parameters
    ----------
    wait: selenium.webdriver.support.wait.WebDriverWait
        Tells the driver to look for an element every 0.5 seconds until it is found

    Returns
    -------
    grid_canvas_right: selenium.webdriver.remote.webelement.WebElement
    """

    logger = logging.getLogger("_locate_grid_canvas_right")

    logger.debug("locate grid canvas right")
    # locate grid canvas right
    grid_canvas_right = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div[class='grid-canvas grid-canvas-top grid-canvas-right']")
        )
    )

    logger.debug("return grid_canvas_right")
    return grid_canvas_right


def _click_enter(driver):
    """
    Tells driver to click ENTER key

    Parameters
    ----------
    driver: selenium.webdriver.chrome.webdriver.WebDriver
        The driver used to control the browser

    Returns
    -------
    None
    """

    logger = logging.getLogger("_click_enter")

    logger.debug("define action chain")
    # define action chain
    actions = ActionChains(driver=driver)

    logger.debug("click ENTER")
    # click ENTER
    actions.send_keys(Keys.ENTER)
    actions.perform()

    logger.debug("return None")
    return None


def enter_status_flag(wait, driver):
    """
    Locates "Enter Status Flag" and changes it from "No" to "Yes"

    Parameters
    ----------
    wait: selenium.webdriver.support.wait.WebDriverWait
        Tells the driver to look for an element every 0.5 seconds until it is found

    driver: selenium.webdriver.chrome.webdriver.WebDriver
        The driver used to control the browser

    Returns
    -------
    None
    """

    logger = logging.getLogger("enter_status_flag")

    logger.debug("_locate_grid_canvas_right")
    grid_canvas_right = _locate_grid_canvas_right(wait=wait)

    logger.debug("locate enter status flag")
    # locate enter status flag
    enter_status_flag_ = grid_canvas_right.find_element_by_css_selector("div")

    logger.debug("find tag with text = No")
    # find tag with text = "No"
    enter_status_flag_ = enter_status_flag_.find_element_by_xpath(
        "//*[text()='No']"
    )

    logger.debug("click it")
    # click it
    enter_status_flag_.click()

    logger.debug("_click_enter")
    _click_enter(driver)

    logger.debug("locate the enter status flag selector")
    # locate the enter status flag selector
    enter_status_flag_selector = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "option[value='Y']")
        )
    )

    logger.debug("click <option>")
    # click <option>
    enter_status_flag_selector.click()

    logger.debug("return None")
    return None


def bi_assignment_owner(wait, driver, bi_assignment_owner_name):
    """
    Locate "BI Assignment Owner" field and select option

    Parameters
    ----------
    wait: selenium.webdriver.support.wait.WebDriverWait
        Tells the driver to look for an element every 0.5 seconds until it is found

    driver: selenium.webdriver.chrome.webdriver.WebDriver
        The driver used to control the browser

    bi_assignment_owner_name: str
        Name of BI Assignment Owner

    Returns
    -------
    None
    """

    logger = logging.getLogger("bi_assignment_owner")

    logger.debug("_locate_grid_canvas_right")
    grid_canvas_right = _locate_grid_canvas_right(wait=wait)

    logger.debug("locate the bi assignment owner")
    # locate the bi assignment owner  # **Analysis and Solution Development**
    bi_assignment_owner_ = grid_canvas_right.find_elements_by_css_selector(
        "div:last-of-type"
        "> div:nth-child(3)"
    )[-1]

    logger.debug("click it")
    # click it
    bi_assignment_owner_.click()

    logger.debug("sleep(1)")
    sleep(1)

    logger.debug("_click_enter")
    _click_enter(driver)

    logger.debug("locate bi assignment owner selector")
    # locate bi assignment owner selector
    bi_assignment_owner_selector = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, f"option[value='{bi_assignment_owner_name}']")
        )
    )

    logger.debug(f"click {bi_assignment_owner_name}")
    # click <option>
    bi_assignment_owner_selector.click()

    logger.debug("return None")
    return None


def bi_team(wait, driver, bi_team_name):
    """
    Locate "BI Team" field and select option

    Parameters
    ----------
    wait: selenium.webdriver.support.wait.WebDriverWait
        Tells the driver to look for an element every 0.5 seconds until it is found

    driver: selenium.webdriver.chrome.webdriver.WebDriver
        The driver used to control the browser

    bi_team_name: str
        Name of BI Team

    Returns
    -------
    None
    """

    logger = logging.getLogger("bi_team")

    logger.debug("_locate_grid_canvas_right")
    grid_canvas_right = _locate_grid_canvas_right(wait)

    logger.debug("locate bi team")
    # locate bi team
    bi_team_ = grid_canvas_right.find_element_by_css_selector(
        "div"
        "> div:nth-child(2)"
    )

    logger.debug("click it")
    # click it
    bi_team_.click()

    logger.debug("sleep(1)")
    sleep(1)

    logger.debug("_click_enter")
    _click_enter(driver)

    logger.debug("locate bi team selector")
    # locate bi team selector
    bi_team_selector = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, f"//*[text()='{bi_team_name}']")
        )
    )

    logger.debug(f"click {bi_team_name}")
    # click <option>
    bi_team_selector.click()

    logger.debug("return None")
    return None


def _locate_grid_canvas_left(wait):
    """
    Locates the grid canvas left and returns it

    Parameters
    ----------
    wait: selenium.webdriver.support.wait.WebDriverWait
        Tells the driver to look for an element every 0.5 seconds until it is found

    Returns
    -------
    grid_canvas_left: selenium.webdriver.remote.webelement.WebElement
    """

    logger = logging.getLogger("_locate_grid_canvas_left")

    logger.debug("locate grid canvas right")
    # locate grid canvas right
    grid_canvas_left = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div[class='grid-canvas grid-canvas-top grid-canvas-left']")
        )
    )

    logger.debug("return grid_canvas_left")
    return grid_canvas_left


def _hover(driver, element):
    """
    Tells driver to hover over element

    Parameters
    ----------
    driver: selenium.webdriver.chrome.webdriver.WebDriver
        The driver used to control the browser

    element: selenium.webdriver.remote.webelement.WebElement
        Element to hover over

    Returns
    -------
    None
    """

    logger = logging.getLogger("_hover")

    logger.debug("define action chain")
    # define action chain
    actions = ActionChains(driver=driver)

    logger.debug("hover over assignments button")
    # hover over assignments button
    actions.move_to_element(element)
    actions.perform()

    logger.debug("return None")
    return None


def open_resource_search_window(wait, driver):
    """
    Open Resource Search window

    Parameters
    ----------
    wait: selenium.webdriver.support.wait.WebDriverWait
        Tells the driver to look for an element every 0.5 seconds until it is found

    driver: selenium.webdriver.chrome.webdriver.WebDriver
        The driver used to control the browser

    Returns
    -------
    None
    """

    logger = logging.getLogger("open_resource_search_window")

    logger.debug("_locate_grid_canvas_left")
    grid_canvas_left = _locate_grid_canvas_left(wait)

    logger.debug("locate analysis and solution development action button")
    # locate analysis and solution development action button
    analysis_and_solution_development_action_button = grid_canvas_left.find_element_by_css_selector(
        "div:last-of-type"
        "> div"
        "> div[class='ActionLinkButton'][title='Actions']"
    )

    logger.debug("click button")
    # click button
    analysis_and_solution_development_action_button.click()

    logger.debug("locate assignments button")
    # locate assignments button
    assignments_button = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//*[text()='Assignments']")
        )
    )

    logger.debug("_hover")
    _hover(driver, assignments_button)

    logger.debug("locate new allocation button")
    # locate new allocation button
    new_allocation_button = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//*[text()='New Allocation']")
        )
    )

    logger.debug("click it")
    # click it
    new_allocation_button.click()

    logger.debug("return None")
    return None


def allocate(driver, wait, resource):
    """
    Allocate a resource.

    Parameters
    ----------
    driver: selenium.webdriver.chrome.webdriver.WebDriver
        The driver used to control the browser

    wait: selenium.webdriver.support.wait.WebDriverWait
        Tells the driver to look for an element every 0.5 seconds until it is found

    resource: str
        Resource to allocate

    Returns
    -------
    None
    """

    logger = logging.getLogger("allocate")

    logger.debug("get resource search window handle")
    # get resource search window handle
    resource_search_window = driver.window_handles[-1]

    logger.debug("switch to resource window")
    # switch to resource window
    driver.switch_to.window(resource_search_window)

    logger.debug("locate search view frame")
    # locate search view frame
    search_view_frame = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "iframe[name='iframeSearchView']")
        )
    )

    logger.debug("switch to search view frame")
    # switch to search view frame
    driver.switch_to.frame(search_view_frame)

    logger.debug("locate search attributes frame")
    # locate search attributes frame
    search_attributes_frame = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "frame[id='frameAttributes']")
        )
    )

    logger.debug("switch to search attributes frame")
    # switch to search attributes frame
    driver.switch_to.frame(search_attributes_frame)

    logger.debug("locate description input")
    # locate description input
    description_input = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[id='attribute_description']")
        )
    )

    logger.debug(f"type {resource}")
    # type <resource name>
    description_input.send_keys(f"{resource}")

    logger.debug("locate search button")
    # locate search button
    search_button = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[name='_search'][type='submit']")
        )
    )

    logger.debug("click button")
    # click button
    search_button.click()

    logger.debug("switch to previous frame")
    # switch to previous frame
    driver.switch_to.parent_frame()

    logger.debug("locate search list frame")
    # locate search list frame
    search_list_frame = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "frame[id='frameSearchList']")
        )
    )

    logger.debug("switch to search list frame")
    # switch to search list frame
    driver.switch_to.frame(search_list_frame)

    logger.debug("locate checkbox input")
    # locate checkbox input
    checkbox_input = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[type='checkbox'][name='sel_list']")
        )
    )

    logger.debug("click checkbox")
    # click checkbox
    checkbox_input.click()

    logger.debug("switch to previous frame")
    # switch to previous frame
    driver.switch_to.parent_frame()

    logger.debug("switch to previous frame (i.e. outermost frame)")
    # switch to previous frame (i.e. outermost frame)
    driver.switch_to.parent_frame()

    logger.debug("locate OK button")
    # locate OK button
    ok_button = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[type='button'][value='OK']")
        )
    )

    logger.debug("click OK button")
    # click OK button
    ok_button.click()

    logger.debug("get main window handle")
    # switch back to main window
    main_window = driver.window_handles[0]

    logger.debug("switch to main window")
    # switch to main window
    driver.switch_to.window(main_window)

    logger.debug("return None")
    return None


def navigate_to_work_view(wait):
    """"""

    #
    action_menu_button = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "span[class='pv12MenuAffordanceIcon'][title='Actions']")
        )
    )

    #
    action_menu_button.click()

    #
    work_view_button = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "li[title='Work View']")
        )
    )

    #
    work_view_button.click()

    return None


def edit_work_detail(driver, wait):
    """"""

    #
    describe_and_categorize_tab = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//*[text()='Describe & Categorize BI']")
        )
    )

    #
    describe_and_categorize_tab.click()

    sleep(2)

    #
    iframe = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "iframe[name='pv-iframeSets-ConfiguredScreens59']")
        )
    )

    #
    driver.switch_to.frame(iframe)

    #
    edit_button = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//*[text()='Edit']")
        )
    )

    #
    edit_button.click()

    return None


def bi_swim_lane(wait, bi_swim_lane_name):
    """"""

    #
    bi_swim_lanes_label = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//*[text()='BI Swim Lanes']")
        )
    )

    #
    bi_swim_lanes_selector = bi_swim_lanes_label.parent.find_element_by_css_selector(
        f"option[value='{bi_swim_lane_name}']")

    #
    bi_swim_lanes_selector.click()

    return None


def bi_work_type(wait):
    """"""

    #
    bi_work_type_selector = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//label[text()='BI Work Type']/parent::div//select/option[@value='3036593']")
        )
    )

    #
    bi_work_type_selector.click()

    return None


def executive_sponsor(wait, executive_sponsor_name):
    """"""

    #
    executive_sponsor_input = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//label[text()='Executive Sponsor']/parent::div//input")
        )
    )

    sleep(1)  # slow down

    #
    # executive_sponsor_input.click()  # not interactable
    executive_sponsor_input.clear()

    #
    executive_sponsor_input.send_keys(executive_sponsor_name)

    return None


def bi_business_owner(wait, bi_business_owner_name):
    """"""

    #
    bi_business_owner_input = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//label[text()='BI Business Owner']/parent::div//input")
        )
    )

    sleep(1)  # slow down

    #
    # bi_business_owner_input.click()  # not interactable
    bi_business_owner_input.clear()

    #
    bi_business_owner_input.send_keys(bi_business_owner_name)

    return None


def bi_domain(wait, driver, bi_domain_name):
    """"""

    #
    bi_domain_input = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "//label[text()='BI Domain']/parent::div//input[@type='text']"
            )
        )
    )

    sleep(1)  # slow down

    #
    # bi_domain_input.click()  # not interactable?
    # bi_domain_input.clear()

    actions = ActionChains(driver=driver)

    actions.move_to_element(bi_domain_input)
    actions.click()
    actions.send_keys(bi_domain_name)
    actions.perform()

    #
    # bi_domain_input.send_keys(bi_domain_name)

    return None


def requestor(wait, driver, requestor_name):
    """"""

    #
    requestor_input = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//label[text()='Requestor']/parent::div//input[@type='text']")
        )
    )

    sleep(1)  # slow down

    actions = ActionChains(driver=driver)

    actions.move_to_element(requestor_input)
    actions.click()
    actions.send_keys(requestor_name)
    actions.perform()

    return None


def bi_liaison(wait, driver, bi_liaison_name):
    """"""

    #
    bi_liaison_select = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, f"//label[text()='BI Liason']/parent::div//select")
        )
    )

    actions = ActionChains(driver=driver)

    actions.click(bi_liaison_select)
    actions.perform()

    bi_liaison_click = bi_liaison_select.find_element_by_css_selector(f"option[value='{bi_liaison_name}']")

    #
    bi_liaison_click.click()

    return None


def work_description(wait, driver, work_description_text):
    """"""

    #
    work_description_text_area = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                f"//div[@title='Detailed Work Description.']"
                "//div[@class='CodeMirror cm-s-paper CodeMirror-wrap']"
                "//div[@class='CodeMirror-lines']"
            )
        )
    )

    #
    work_description_text_area.click()

    #
    actions = ActionChains(driver=driver)

    #
    actions.send_keys(work_description_text)
    actions.perform()

    return None


def business_need(wait, driver, business_need_text):
    """"""

    #
    business_need_text_area = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                f"//label[text()='Business Need']/parent::div/div[@class='attribute-field']"
                "//div[@class='CodeMirror cm-s-paper CodeMirror-wrap']"
                "//div[@class='CodeMirror-lines']"
            )
        )
    )

    #
    business_need_text_area.click()

    #
    actions = ActionChains(driver=driver)

    #
    actions.send_keys(business_need_text)
    actions.perform()

    return None


def save_edits(wait):
    """"""

    save_button = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//button/span/span[@class='button-text'][text()='Save']")
        )
    )

    save_button.click()

    return None


if __name__ == "__main__":

    logger = logging.getLogger(__name__)

    log_date = str(pd.to_datetime("today").date()).replace("-", "")

    logging.basicConfig(
        filename=f"Logs/PRM_{log_date}.log",
        filemode="w",
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    # adding StreamHandler so logs print to terminal
    logger.addHandler(logging.StreamHandler())

    logger.info("read data")
    # read data
    df = pd.read_excel(
        io=r"\\kuha.kumed.com\shares\departments\AA & BIS\PRM\2022494_Automation_of_PRM_Task_Creation\ProjectBook.xlsx"
        # io=r"\\kuha.kumed.com\shares\departments\AA & BIS\PRM\2022494_Automation_of_PRM_Task_Creation\ProjectBookTest.xlsx"
    )

    logger.info("define options")
    # define options
    options = Options()
    # options.headless = True

    logger.info("_get_date")
    today_str = _get_date()
    date = pd.to_datetime(today_str)
    month = date.month_name()
    year = date.year

    logger.info("create driver")
    # create driver
    driver = webdriver.Chrome(options=options)

    logger.info("define wait")
    # define wait
    wait = WebDriverWait(driver=driver, timeout=30)

    for row in df.itertuples():

        logger.info("get project info")
        description = row.Description
        description = f"{month} {year} {description}"

        bi_service_name = row.BIServiceName

        bi_assignment_owner_name = row.BIAssignmentOwner
        left = bi_assignment_owner_name.index("(")
        right = bi_assignment_owner_name.index(")")
        bi_assignment_owner_name = bi_assignment_owner_name[left + 1: right]

        bi_team_name = row.BITeam

        resource = row.BIAssignmentOwner
        left = resource.index("(")
        resource = resource[:left]
        resource = resource.strip()

        left = row.BISwimLanes.index("(")
        bi_swim_lane_name = row.BISwimLanes[left + 1:-1]

        executive_sponsor_name = row.ExecutiveSponsor

        bi_business_owner_name = row.BIBusinessOwner

        bi_domain_name = row.BIDomain

        requestor_name = row.Requestor

        bi_liaison_name = row.BILiaison
        left = bi_liaison_name.index("(")
        bi_liaison_name = bi_liaison_name[left + 1: -1]

        work_description_text = row.WorkDescription

        business_need_text = row.BusinessNeed

        logger.info("login")
        login(driver, wait)

        logger.info(f"new_work_page1 {description}")
        new_work_page1(wait, today_str, description)

        logger.info("new_work_page2")
        new_work_page2(wait, today_str, bi_service_name)

        logger.info("enter_status_flag")
        enter_status_flag(wait, driver)

        logger.info("bi_assignment_owner")
        bi_assignment_owner(wait, driver, bi_assignment_owner_name)

        logger.info("bi_team")
        bi_team(wait, driver, bi_team_name)

        logger.info("open_resource_search_window")
        open_resource_search_window(wait, driver)

        logger.info("allocate")
        allocate(driver, wait, resource)

        sleep(2)

        logger.info("navigate_to_work_view")
        navigate_to_work_view(wait)

        logger.info("edit_work_detail")
        edit_work_detail(driver, wait)

        logger.info("bi_swim_lane")
        bi_swim_lane(wait, bi_swim_lane_name)

        logger.info("bi_work_type")
        bi_work_type(wait)

        logger.info("executive_sponsor")
        executive_sponsor(wait, executive_sponsor_name)

        logger.info("bi_business_owner")
        bi_business_owner(wait, bi_business_owner_name)

        logger.info("bi_domain")
        bi_domain(wait, driver, bi_domain_name)

        logger.info("requestor")
        requestor(wait, driver, requestor_name)

        logger.info("bi_liaison")
        bi_liaison(wait, driver, bi_liaison_name)

        logger.info("work_description")
        work_description(wait, driver, work_description_text)

        logger.info("business_need")
        business_need(wait, driver, business_need_text)

        logger.info("save_edits")
        save_edits(wait)

        sleep(2)

    logger.info("quit")
    driver.quit()
