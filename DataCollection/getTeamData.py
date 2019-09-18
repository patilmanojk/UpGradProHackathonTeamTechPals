from selenium import webdriver

from Utils.ExelUtil import ExcelUtil

driver = webdriver.Firefox()
driver.get("https://www.prokabaddi.com")
driver.implicitly_wait(30)

driver.find_element_by_xpath("//li[@data-id = 'teams']").click()

teams = driver.find_elements_by_xpath("//li[@data-id = 'teams']/div/a")
teamnames = []
for team_name in teams:
    teamnames.append(team_name.text)


for team in teamnames:
    TeamName = team
    print(TeamName)

    driver.find_element_by_xpath("// a[ @ href = '/']").click()
    driver.find_element_by_xpath("//li[@data-id = 'teams']").click()

    driver.find_element_by_link_text(TeamName).click()
    ExcelUtil.addSheet("ProKabbaddi.xlsx",TeamName)
    position_cat_name = []
    position_cat = driver.find_elements_by_xpath("//div[@class = 'si-section-header']")
    for position in position_cat:
        position_cat_name.append(position.text)

    for i in position_cat_name:
        if i == 'OVERALL':
            ExcelUtil.writedatatoMultipleCells("ProKabbaddi.xlsx", TeamName,"A",1,8,i)
        if i == 'ATTACK':
            ExcelUtil.writedatatoMultipleCells("ProKabbaddi.xlsx", TeamName,"A",8,17,i)
        if i == 'DEFENCE':
            ExcelUtil.writedatatoMultipleCells("ProKabbaddi.xlsx", TeamName,"A",17,25,i)

        start_pos = 1
        end_pos = 1

    for i in range(2,5):
        Category_name = driver.find_element_by_xpath("(//div[@class = 'si-tbl-data'])["+str(i)+"]").text.splitlines()
        cat_len = len(Category_name)
        end_pos += cat_len
        for j in Category_name:
            ExcelUtil.writedataSingleCell("ProKabbaddi.xlsx", TeamName, "B", start_pos , j)
            start_pos += 1
    Season_name = driver.find_elements_by_xpath("//div[@class = 'si-stats-container']/div[2]/div[1]/div[@class = 'si-tbl-data']")
    cat_len = len(Season_name)
    cell_pos = 2
    cell_col = [chr(i) for i in range(ord('A'),ord('Z')+1)]
    for j in Season_name:
        ExcelUtil.writedataSingleCell("ProKabbaddi.xlsx",TeamName, cell_col[cell_pos], 0 , j.text)
        for k in range(1,6):
            cell_value = driver.find_element_by_xpath("//div[@class = 'si-stats-container']/div[2]/div[2]/div["+str(cell_pos-1)+"]/div["+str(k)+"]")
            ExcelUtil.writedataSingleCell("ProKabbaddi.xlsx", TeamName, cell_col[cell_pos], k, cell_value.text)
        for l in range(1, 10):
            cell_value = driver.find_element_by_xpath("//div[@class = 'si-stats-container']/div[2]/div[3]/div["+str(cell_pos-1)+"]/div["+str(l)+"]")
            ExcelUtil.writedataSingleCell("ProKabbaddi.xlsx", TeamName, cell_col[cell_pos], 7+l, cell_value.text)
        for m in range(1, 9):
            cell_value = driver.find_element_by_xpath("//div[@class = 'si-stats-container']/div[2]/div[4]/div["+str(cell_pos-1)+"]/div["+str(m)+"]")
            ExcelUtil.writedataSingleCell("ProKabbaddi.xlsx", TeamName, cell_col[cell_pos], 16+m, cell_value.text)
        cell_pos += 1


driver.quit()
