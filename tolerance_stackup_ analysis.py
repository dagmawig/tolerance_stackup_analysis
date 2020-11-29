# this python program is used to do monte carlo tolerance stack up analysis of dimensions
# given the dimensions and their tolerances and the direction of each dimension
# this program simulates 10,000 variation of each dimensions and calculates the sum total/gap 
# for each variation. The variation is simulated using gaussian distribution of each 
# dimension within the given tolerance where the mean is halfway between the upper and lower
# spec limit of each dimension and the standard deviation is one sixth of the difference 
# between the upper and lower spec limit

# importing all the required modules
import xlrd
from array import array
from random import normalvariate
import statistics as stat
import matplotlib.pyplot as plt

# file location/ example file included in the project folder
loc = ('C:/Users/Dag/Documents/webdev/python/monte_carlo/raw.xlsx')

# open an excel file
raw = xlrd.open_workbook('raw1.xlsx')

# pulling the first sheet of the file
sheet = raw.sheet_by_index(0)


# define nominal values and tolerances and dimension signs
def nominal(i):
    return sheet.cell_value(i+1, 1)
def upperLim(i):
    return nominal(i) + sheet.cell_value(i+1, 2)
def lowerLim(i):
    return nominal(i) - sheet.cell_value(i+1, 3)
def dimSign(i):
    return sheet.cell_value(i+1, 4)
def sigma(i):
    return ( upperLim(i) - lowerLim(i) ) / 6
def mean(i):
    return ( upperLim(i) + lowerLim(i) ) / 2

# calculating the number of dimensions given
count = len(sheet.col_values(0)) - 1

# defining the number of randomly generated variations for a given dimension
# the higher this number the better the analysis
n = 10000

# define an array to contain all vairations of all dimensions
dimArray = [ [ None ] * n for i in range(count) ]

# populating the array with randomly generated variation of the dimensions
for i in range(count):
    for j in range(n):
        aver = mean(i)
        std = sigma(i)
        dimArray[i][j] =   normalvariate(aver, std)

# defining an array containing sum total of dimensions for each of the n vairations
gap = [ 0.0 ] * n

# populating the gap array
for j in range(n):
    for i in range(count):
        gap[j] = gap[j] + ( dimSign(i) * dimArray[i][j] )

# plotting the gap vairation in a histogram chart to see how many of the variations fit
# within the allowable gap range 
average = stat.mean(gap)
stdDev = stat.stdev(gap)
minimum = min(gap)-stdDev
maximum = max(gap)+stdDev
rangeOfVal = ( minimum, maximum )
bins = 8
ticksArr = [minimum + (maximum-minimum)/(2*bins) + i*(maximum - minimum)/bins for i in range(bins)]
graph = plt.hist(gap , bins, rangeOfVal, color = 'green', histtype = 'bar', rwidth = .8)
plt.xlabel('gap')
plt.ylabel('frequency')
plt.title('gap distribution')
plt.xticks(ticksArr)
for i in range(bins):
    plt.text(graph[1][i] + (maximum-minimum)*.1/(bins) , graph[0][i], str(graph[0][i]))

# showing the histogram plot
plt.show()
