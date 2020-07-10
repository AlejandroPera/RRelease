import asyncio
from pyppeteer import launch

RATE_RECORD_MOD = "https://dsctmststr2.dhl.com/GC3/glog.webserver.finder.FinderServlet?ct=MTgxMzA2MDUwOTgwMTEwNTA2NQ%3D%3D&query_name=glog.server.query.rate.RateGeoQuery&finder_set_gid=MXPLAN.MX%20RECORD"
RATE_OFFERING_MOD = "https://dsctmststr2.dhl.com/GC3/glog.webserver.finder.FinderServlet?ct=NzExMTkyNjc2NjMzNTM4ODk3OQ%3D%3D&query_name=glog.server.query.rate.RateOfferingQuery&finder_set_gid=MXPLAN.MX%20OFFERING"
SAMPLE_RID = "d121F_1927635_48M-CON_CHI_TAPACHULA"

#Function for reading restricted local credentials
def readLoginCred(filename):
    f = open(filename, "r")
    user = f.readline().strip()
    psswd = f.readline()
    output = [user,psswd]
    return output

async def rateOffering(page):
    await page.goto(RATE_OFFERING_MOD)
    return 1

async def rateRecord(page):
    await page.goto("https://dsctmststr2.dhl.com/GC3/glog.webserver.servlet.umt.Login")
    username_field = await page.J("body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]")
    password_field = await page.J("body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > form > table > tbody > tr:nth-child(3) > td:nth-child(2) > input[type=password]")
    credentials = readLoginCred("cred.txt")
    await username_field.type(credentials[0])
    await password_field.type(credentials[1])
    await page.waitForNavigation()

    #2
    await page.goto(RATE_RECORD_MOD)
    record_field = await page.J("#bodyDataDiv > table > tbody > tr > td:nth-child(1) > table > tbody > tr:nth-child(1) > td > div > input[type=text]:nth-child(3)")
    await record_field.type(SAMPLE_RID)
    isActive_drop = await page.J("#bodyDataDiv > table > tbody > tr > td:nth-child(1) > table > tbody > tr:nth-child(2) > td > div > select > option:nth-child(2)")
    search_button = await page.J("#search_button")
    await isActive_drop.click()
    await search_button.click()

    #3
    await page.waitForNavigation()
    await page.waitFor("#resultsPage\:pgl2 > tbody > tr > td > div > table > tbody > tr > td:nth-child(2) > span:nth-child(2)")
    element = await page.J("#resultsPage\:pgl2 > tbody > tr > td > div > table > tbody > tr > td:nth-child(2) > span:nth-child(2)")
    #Extracts innerText of element
    total_results = int (await page.evaluate('(element) => element.innerText', element))
    if total_results > 0:
        cb = await page.J("#rgSGSec\.1\.1\.1\.1\.check")
        cb.click()
    else:
        await rateOffering(page)        

    return 1

#Scenario 1 Rate already exists
async def rateRecordOTM():
    browser = await launch(headless=False)  # headless false means open the browser in the operation
    page = await browser.newPage()
    await page.setViewport({"width": 1024, "height": 768, "deviceScaleFactor": 1})
    page.setDefaultNavigationTimeout(60000)
    await page.goto("https://dsctmststr2.dhl.com/GC3/glog.webserver.servlet.umt.Login")
    credentials = readLoginCred("cred.txt")
    passw=await page.waitFor("[name='userpassword']")
    usernN=await page.waitFor("[name='username']")
    await usernN.type(credential[0])
    await passw.type(credential[1])  
    await page.click("[name='submitbutton']")   
    
    #await page.waitForNavigation()
    await page.goto("https://dsctmststr2.dhl.com/GC3/glog.webserver.finder.FinderServlet?ct=MTgxMzA2MDUwOTgwMTEwNTA2NQ%3D%3D&query_name=glog.server.query.rate.RateGeoQuery&finder_set_gid=MXPLAN.MX%20RECORD")
    await page.type("#bodyDataDiv > table > tbody > tr > td:nth-child(1) > table > tbody > tr:nth-child(1) > td > div > input[type=text]:nth-child(3)",SAMPLE_RID)
    await page.click("#bodyDataDiv > table > tbody > tr > td:nth-child(1) > table > tbody > tr:nth-child(2) > td > div > select > option:nth-child(2)")
    await page.click("#search_button")
    
   
    await page.waitForNavigation()
    await page.waitFor("#resultsPage\:pgl2 > tbody > tr > td > div > table > tbody > tr > td:nth-child(2) > span:nth-child(2)")
    element = await page.J("#resultsPage\:pgl2 > tbody > tr > td > div > table > tbody > tr > td:nth-child(2) > span:nth-child(2)")
    #Extracts innerText of element
    total_results = int (await page.evaluate('(element) => element.innerText', element))
    if total_results > 0:
        cb = await page.J("#rgSGSec\.1\.1\.1\.1\.check")
        await cb.click()
    else:
        await rateOffering(page)
      
    await page.click("#viewButton")
    
    #Locate and manage the corresping pop up
    popI = len(browser.targets())
    print(popI)
    pop = browser.targets()[popI-1]
    popup = await pop.page()
    #offeringIDElement = await popup.J("#bodyDataDiv > div:nth-child(1) > table > tbody > tr:nth-child(1) > td:nth-child(1) > div > div.fieldData")
    print(popup.url)
    offeringID = await page.evaluate('(element) => element.innerText',"#bodyDataDiv > div:nth-child(1) > table > tbody > tr:nth-child(1) > td:nth-child(1) > div > div.fieldData")
    print(offeringID)
    await page.waitFor(6000) 
    return 1


async def main():
    browser = await launch(headless = False)
    page = await browser.newPage()
    #Steps 1-2
    await rateRecord(page)
    await page.waitFor(6000)
    await browser.close()

asyncio.get_event_loop().run_until_complete(main())

