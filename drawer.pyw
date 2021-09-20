from os import mkdir
from os.path import exists
from numpy import corrcoef, array
from pandas import read_excel
from matplotlib import rc
from matplotlib.font_manager import FontProperties
from matplotlib.pyplot import figure, plot, subplot, pie, scatter, legend, title, close, savefig


'''
    OSMU: All In One

    TODO: Express revenue and profit in line charts
    TODO: Changes in Product composition (pie charts)
    TODO: COVID 19 Impact by Product

    Excel DataFrame Size: 16Rx13C
    Columns: Unnamed index;

    Cosmetic: 화장품
    Toothpaste: 치약
    Moistrure: 보습제


'''

CONSTANT_MATPLOTLIB_HANGEUL = "C:/Windows/Fonts/gulim.ttc"
CONSTANT_RESULT_EXPORT_PATH = "./exports"
CONSTANT_FILE_PATH = "access.xlsx"
CONSTANT_SHEET_NAME = "company_sales_data"


dataframe = read_excel(CONSTANT_FILE_PATH, CONSTANT_SHEET_NAME)

# 0. Enum

KnownColumnIndex = {
    "Common": {
        "Revenue": {
            "cosmetic": "Unnamed: 5",
            "toothpaste": "Unnamed: 6",
            "moisture": "Unnamed: 7",
            "Total": "Unnamed: 8"
        },
        "Profilt": {
            "cosmetic": "Unnamed: 9",
            "toothpaste": "Unnamed: 10",
            "moisture": "Unnamed: 11",
            "Total": "Unnamed: 12"
        },
        "Realm": {
            "EconomicIndex": "Unnamed: 3",
            "CovidCases": "Unnamed: 4"
        }
    },
    "Quarter": {
        "2018": [2, 3, 4, 5],
        "2019": [6, 7, 8, 9],
        "2020": [10, 11, 12, 13],
        "2021": [14, 15]
    }
}

# 1. Line Charts
def draw_line_charts():
    seed = dataframe.loc[2: 15]

    figure(figsize=(15, 6))

    subplot(1, 2, 1)

    plot(seed[KnownColumnIndex["Common"]["Revenue"]["cosmetic"]])
    plot(seed[KnownColumnIndex["Common"]["Revenue"]["toothpaste"]])
    plot(seed[KnownColumnIndex["Common"]["Revenue"]["moisture"]])
    plot(seed[KnownColumnIndex["Common"]["Revenue"]["Total"]])

    title('revenue')

    legend(KnownColumnIndex["Common"]["Revenue"].keys())

    subplot(1, 2, 2)

    plot(seed[KnownColumnIndex["Common"]["Profilt"]["cosmetic"]])
    plot(seed[KnownColumnIndex["Common"]["Profilt"]["toothpaste"]])
    plot(seed[KnownColumnIndex["Common"]["Profilt"]["moisture"]])
    plot(seed[KnownColumnIndex["Common"]["Profilt"]["Total"]])

    title('profilt')

    legend(KnownColumnIndex["Common"]["Profilt"].keys())
    savefig('{}/LineChart/LineChart.png'.format(CONSTANT_RESULT_EXPORT_PATH))
    
    close('all')

# 2. Pie Charts
def draw_pie_chart():
    # 제품 구성의 변경사항?? 무슨 뜻이지? 몰라, 걍 다 만들어...;;
    unit = "분기"
    def get_values_with_seed(col, row):
        return dataframe.loc[col[0]: col[-1], row].values, [str(quarter+1)+unit for quarter in range(len(col))]
    
    def def_for_quick_draw(Year, Group, Product):
        data, label = get_values_with_seed(KnownColumnIndex["Quarter"][Year], KnownColumnIndex["Common"][Group][Product])
        pie(data, labels=label)
        title('{} of {} in {}'.format(Group, Product, Year))

        # Make Directory

        if not exists(CONSTANT_RESULT_EXPORT_PATH+"/PieChart/{}".format(Group)):
            mkdir(CONSTANT_RESULT_EXPORT_PATH+"/PieChart/{}".format(Group))

        if not exists(CONSTANT_RESULT_EXPORT_PATH+"/PieChart/{}/{}".format(Group, Year)):
            mkdir(CONSTANT_RESULT_EXPORT_PATH+"/PieChart/{}/{}".format(Group, Year))
        
        # End Make Directory
        
        savefig('{}/PieChart/{}/{}/{}.png'.format(CONSTANT_RESULT_EXPORT_PATH, Group, Year, Product))
        close('all')

    ''''''''''''''''''''''''''''''''''''
    # 각 상품들의 분기별 매출에 대하여...
    ''''''''''''''''''''''''''''''''''''
    # 1. Cosmetic
    def_for_quick_draw("2018", "Revenue", "cosmetic")
    def_for_quick_draw("2019", "Revenue", "cosmetic")
    def_for_quick_draw("2020", "Revenue", "cosmetic")
    def_for_quick_draw("2021", "Revenue", "cosmetic")

    # 2. Toothpaste
    def_for_quick_draw("2018", "Revenue", "toothpaste")
    def_for_quick_draw("2019", "Revenue", "toothpaste")
    def_for_quick_draw("2020", "Revenue", "toothpaste")
    def_for_quick_draw("2021", "Revenue", "toothpaste")

    # 3. Moisture
    def_for_quick_draw("2018", "Revenue", "moisture")
    def_for_quick_draw("2019", "Revenue", "moisture")
    def_for_quick_draw("2020", "Revenue", "moisture")
    def_for_quick_draw("2021", "Revenue", "moisture")

    # 4. Total
    def_for_quick_draw("2018", "Revenue", "Total")
    def_for_quick_draw("2019", "Revenue", "Total")
    def_for_quick_draw("2020", "Revenue", "Total")
    def_for_quick_draw("2021", "Revenue", "Total")

    ''''''''''''''''''''''''''''''''''''
    # 각 상품들의 분기별 순 이익에 대하여...
    ''''''''''''''''''''''''''''''''''''
    # 1. Cosmetic
    def_for_quick_draw("2018", "Profilt", "cosmetic")
    def_for_quick_draw("2019", "Profilt", "cosmetic")
    def_for_quick_draw("2020", "Profilt", "cosmetic")
    def_for_quick_draw("2021", "Profilt", "cosmetic")

    # 2. Toothpaste
    def_for_quick_draw("2018", "Profilt", "toothpaste")
    def_for_quick_draw("2019", "Profilt", "toothpaste")
    def_for_quick_draw("2020", "Profilt", "toothpaste")
    def_for_quick_draw("2021", "Profilt", "toothpaste")

    # 3. Moisture
    def_for_quick_draw("2018", "Profilt", "moisture")
    def_for_quick_draw("2019", "Profilt", "moisture")
    def_for_quick_draw("2020", "Profilt", "moisture")
    def_for_quick_draw("2021", "Profilt", "moisture")

    # 4. Total
    def_for_quick_draw("2018", "Profilt", "Total")
    def_for_quick_draw("2019", "Profilt", "Total")
    def_for_quick_draw("2020", "Profilt", "Total")
    def_for_quick_draw("2021", "Profilt", "Total")


# 3. Scatter Plot
def draw_scatter_plot():
    # 상관계수: 두 변수 X, Y 사이의 상관계수를 나타낸 수치
    seed = dataframe.loc[2: 15]
    def def_for_quick_draw(SalesTypeX: str, ProductX: str, SalesTypeY: str, ProductY: str) -> None:
        X = seed[KnownColumnIndex["Common"][SalesTypeX][ProductX]]
        Y = seed[KnownColumnIndex["Common"][SalesTypeY][ProductY]]
        scatter(X, Y)
        
        # End Make Directory

        # get_correlation_coefficient
        title("correlation_coefficient: {}".format(corrcoef(array(X).tolist(), array(Y).tolist())[0, 1]))
        savefig('{}/ScatterPlot/{}-{}.png'.format(CONSTANT_RESULT_EXPORT_PATH, ProductX+"_"+SalesTypeX, ProductY+"_"+SalesTypeY))
        close('all')

    ''''''''''''''''''''''''''''''''''''
    # 매출과 순 이익사이의 상관계수와 그 산점도에 대하여...
    ''''''''''''''''''''''''''''''''''''
    def_for_quick_draw("Revenue", "cosmetic", "Profilt", "cosmetic")
    def_for_quick_draw("Revenue", "toothpaste", "Profilt", "toothpaste")
    def_for_quick_draw("Revenue", "moisture", "Profilt", "moisture")
    def_for_quick_draw("Revenue", "Total", "Profilt", "Total")

    ''''''''''''''''''''''''''''''''''''
    # 코로나 확진자수의 증가와 매출-이익 사이의 상관계수와 산점도에 대하여...
    ''''''''''''''''''''''''''''''''''''
    # 매출
    def_for_quick_draw("Realm", "CovidCases", "Revenue", "cosmetic")
    def_for_quick_draw("Realm", "CovidCases", "Revenue", "toothpaste")
    def_for_quick_draw("Realm", "CovidCases", "Revenue", "moisture")
    def_for_quick_draw("Realm", "CovidCases", "Revenue", "Total")

    # 순 이익
    def_for_quick_draw("Realm", "CovidCases", "Profilt", "cosmetic")
    def_for_quick_draw("Realm", "CovidCases", "Profilt", "toothpaste")
    def_for_quick_draw("Realm", "CovidCases", "Profilt", "moisture")
    def_for_quick_draw("Realm", "CovidCases", "Profilt", "Total")

    ''''''''''''''''''''''''''''''''''''
    # 경제지수와 매출-이익 사이의 상관계수와 그 산점도에 대하여...
    ''''''''''''''''''''''''''''''''''''
    # 매출
    def_for_quick_draw("Realm", "EconomicIndex", "Revenue", "cosmetic")
    def_for_quick_draw("Realm", "EconomicIndex", "Revenue", "toothpaste")
    def_for_quick_draw("Realm", "EconomicIndex", "Revenue", "moisture")
    def_for_quick_draw("Realm", "EconomicIndex", "Revenue", "Total")

    # 순 이익
    def_for_quick_draw("Realm", "EconomicIndex", "Profilt", "cosmetic")
    def_for_quick_draw("Realm", "EconomicIndex", "Profilt", "toothpaste")
    def_for_quick_draw("Realm", "EconomicIndex", "Profilt", "moisture")
    def_for_quick_draw("Realm", "EconomicIndex", "Profilt", "Total")

    ''''''''''''''''''''''''''''''''''''
    # 코로나 확진자수의 증가와 경제지수 사이의 상관계수와 산점도에 대하여...
    ''''''''''''''''''''''''''''''''''''
    def_for_quick_draw("Realm", "EconomicIndex", "Realm", "CovidCases")

# 4. Initalize to Export Charts as Images
if __name__ == "__main__":
    rc('font', family=FontProperties(fname=CONSTANT_MATPLOTLIB_HANGEUL).get_name())

    # Make Directory

    if not exists(CONSTANT_RESULT_EXPORT_PATH):
        mkdir(CONSTANT_RESULT_EXPORT_PATH)

    if not exists(CONSTANT_RESULT_EXPORT_PATH+"/LineChart"):
        mkdir(CONSTANT_RESULT_EXPORT_PATH+"/LineChart")

    if not exists(CONSTANT_RESULT_EXPORT_PATH+"/PieChart"):
        mkdir(CONSTANT_RESULT_EXPORT_PATH+"/PieChart")

    if not exists(CONSTANT_RESULT_EXPORT_PATH+"/ScatterPlot"):
        mkdir(CONSTANT_RESULT_EXPORT_PATH+"/ScatterPlot")

    # End Make Directory

    draw_line_charts()
    draw_pie_chart()
    draw_scatter_plot()
    input()
