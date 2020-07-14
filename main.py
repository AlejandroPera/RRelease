import asyncio
from pyppeteer import launch

RATE_RECORD_MOD = "https://dsctmststr2.dhl.com/GC3/glog.webserver.finder.FinderServlet?ct=MTgxMzA2MDUwOTgwMTEwNTA2NQ%3D%3D&query_name=glog.server.query.rate.RateGeoQuery&finder_set_gid=MXPLAN.MX%20RECORD"
RATE_OFFERING_MOD = "https://dsctmststr2.dhl.com/GC3/glog.webserver.finder.FinderServlet?ct=NzExMTkyNjc2NjMzNTM4ODk3OQ%3D%3D&query_name=glog.server.query.rate.RateOfferingQuery&finder_set_gid=MXPLAN.MX%20OFFERING"
SAMPLE_RID = "115F_6049735_THN"

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
    #await page.setViewport({"width": 1024, "height": 768, "deviceScaleFactor": 1})
    page.setDefaultNavigationTimeout(60000)

    #1 Login
    await page.goto("https://dsctmststr2.dhl.com/GC3/glog.webserver.servlet.umt.Login")
    credentials = readLoginCred("cred.txt")
    passw=await page.waitFor("[name='userpassword']")
    usernN=await page.waitFor("[name='username']")
    await usernN.type(credentials[0])
    await passw.type(credentials[1])        
    await page.click("[name='submitbutton']")
    #await page.waitForNavigation()

    #2 Search rate
    await page.goto("https://dsctmststr2.dhl.com/GC3/glog.webserver.finder.FinderServlet?ct=MTgxMzA2MDUwOTgwMTEwNTA2NQ%3D%3D&query_name=glog.server.query.rate.RateGeoQuery&finder_set_gid=MXPLAN.MX%20RECORD")  
    await page.type("#bodyDataDiv > table > tbody > tr > td:nth-child(1) > table > tbody > tr:nth-child(1) > td > div > input[type=text]:nth-child(3)",SAMPLE_RID) 
    await page.click("#bodyDataDiv > table > tbody > tr > td:nth-child(1) > table > tbody > tr:nth-child(2) > td > div > select > option:nth-child(2)")
    await page.click("#search_button")
    await page.waitForNavigation()
    
    #3 Validate every rate via Rate Offering ID
    await page.waitFor("#resultsPage\:pgl2 > tbody > tr > td > div > table > tbody > tr > td:nth-child(2) > span:nth-child(2)")
    element = await page.J("#resultsPage\:pgl2 > tbody > tr > td > div > table > tbody > tr > td:nth-child(2) > span:nth-child(2)")
    #Extracts innerText of element
    total_results = int (await page.evaluate('(element) => element.innerText', element))
    if total_results > 0:
        cb = await page.J("#rgSGSec\.1\.1\.1\.1\.check")
        await cb.click()
        element = await page.J("#rgPageSelectedTotal")
        #Total selected
        selected_total = int (await page.evaluate('(element) => element.innerText', element))
        #Click on the first one
        #bodyDataDiv > div:nth-child(1) > table > tbody > tr:nth-child(1) > td:nth-child(1) > div > div.fieldData
        #for i in range(1,selected_total):
        #await page.click("#rgSGSec\\2e 2\\2e 2\\2e {}\\2e 1 > a".format(1))
        
        await page.goto("https://dsctmststr2.dhl.com/GC3/glog.webserver.finder.WindowOpenFramesetServlet/nss?ct=MTY0NTY0NTk3NDMzMDk2OTczMw%3D%3D&url=glog.webserver.util.ViewDisplayServlet%3Fquery_name%3Dglog.server.query.rate.RateOfferingQuery%26finder_set_gid%3D%26xid%3D115F_6049735_THN%26label%3DRate%20Offering%20ID%26gid%3DMXPLAN.115F_6049735_THN&is_new_window=true")
        body = await page.evaluate('document.body.textContent', force_expr=True)
        print(body)
        frames = len(page.frames)
        print(frames)

        popup = page.frames[0]
        
        '''pages = await browser.pages()
        pagesLen = len(pages)
        #print(pagesLen)
        secondPage = pages[pagesLen-1]
        print(secondPage.url)'''
        '''popLen = len(browser.targets())
        print(popLen)
        pop = browser.targets()[popLen-1]
        #print(pop)
        popup = await pop.page()
        #print(popup.url)
        frame= popup.frames[0]
        print(frame.url)
        #print(type(frame))
        
        
        #4 Validate Rate offering ID
        element = await popup.J("#bodyDataDiv > div:nth-child(1) > table > tbody > tr:nth-child(1) > td:nth-child(1) > div > div.fieldData")
        
        offeringID = str (await popup.evaluate('(element) => element.innerText', element))
        print(offeringID)'''

        
        await page.waitForNavigation() 

    else:
        await rateOffering(page)

    await page.waitForNavigation()


async def main():
    browser = await launch(headless = False)
    page = await browser.newPage()
    #Steps 1-2
    await rateRecord(page)
    await page.waitFor(6000)
    await browser.close()

#asyncio.get_event_loop().run_until_complete(main())
asyncio.get_event_loop().run_until_complete(rateRecordOTM())
