from time import *
import matplotlib.pyplot as plt
import numpy as np
from scipy.stats import *
import xlwings as xw
import random as r

wb = xw.Book('Optimal.xlsx')
WACC = wb.sheets['WACC']
input_sheet = wb.sheets['Cost of Capital Schedule']
sheet4 = wb.sheets['Sheet4']
df_spread = wb.sheets['Default spread']
historical_data = wb.sheets['Historical DAta']

# Bootstrap function
def bootstrap(data,n_samples,statfunc):
    data = np.array(data)
    resampled_stat = []
    for k in range(n_samples):
        index = np.random.randint(0,len(data),len(data))
        sample = data[index]
        bstat = statfunc(sample)
        resampled_stat.append(bstat)
    return np.array(resampled_stat)

cfo_ebit_ratio_data = historical_data.range('F17:F43').value
bootstrap_data = bootstrap(cfo_ebit_ratio_data,100000,np.mean) # 100,000 bootstrap samples
conf_int = np.percentile(bootstrap_data,[0.5,99.5])

# Descriptive Statistics - Bootstrapped data
print('Empirical mean: %.4f' %np.mean(bootstrap_data))
print('99% Confidence interval : ',conf_int)
print('Standard deviation: %.5f'%(np.std(bootstrap_data)))

# Plotting bootstrapped data histogram
plt.hist(bootstrap_data,bins=50,color='g',edgecolor='black',alpha=0.7)
plt.title('EBIT-Cash coverage multiple')
plt.ylabel('Frequency',fontsize=8)
plt.tight_layout()
plt.show()


#Memoization & Iteration_default_spread
cache = {}
def add_spread(rating):
    if rating in cache:
        return cache[rating]
    else:
        for i in range(5,20):
            if rating != df_spread.cells(i,'E').value:
                pass
            else:
                spread = df_spread.cells(i,'F').value
                cache[rating] = spread
                return spread

def change_inputs():
    # Distributions of inputs (Risk-free rate, ERP, CFO-Pre WC)
    global risk_free
    risk_free = r.triangular(0.0375,0.055,0.04) # risk-free rate
    equity_risk_prm = skewnorm.rvs(-3,loc=0.055,scale=0.0055)  # ERP - Source: Kroll report
    ebit_CFO_ratio = r.normalvariate(0.4678,0.06866)  # CFO/EBIT ratio - estimated by bootstrapping
    cfo_pre_wc = skewnorm.rvs(3.5,loc=750,scale=40) #CFO Pre-WC
    input_lst = [risk_free,equity_risk_prm,ebit_CFO_ratio,cfo_pre_wc]
    input_sheet.range('E13').options(transpose=True).value = input_lst

def pre_tax_cod_iteration():
    list_ratings = sheet4.range('E9:M9').value
    list_rate_iter_1 = []
    for rating_1 in list_ratings:
        real_spread = add_spread(rating_1)
        list_rate_iter_1.append(risk_free + real_spread)
    sheet4.range('E10:M10').value = list_rate_iter_1

def expected_value(index,lst): # the index represents debt level - For ex, 2 for 20%
    lst_1 = []
    for i in lst:
        lst_1.append(i[index])
    return np.mean(lst_1) # to get the expected WACC at debt level (index)


lst_outputs = []
#Simulations
for i in range(10000): # Just change the range to number of simulations you would like to have
    change_inputs()
    sheet4.range('E10:M10').options(transpose=True).value = 0.01
    for j in range(3):
        pre_tax_cod_iteration()
    lst_outputs.append(input_sheet.cells(27,'G').value)

plt.hist(lst_outputs,bins=50,color='blue',range=[0.05,0.1], density=True, edgecolor='black')
plt.title('Distribution - WACC at 60% leverage ratio')
plt.ylabel('Frequency',fontsize=8)
plt.tight_layout()
plt.show()
print(np.mean(lst_outputs), np.std(lst_outputs))
print(np.percentile(lst_outputs,[0,10,20,30,40,50,60,70,80,90]))
print(np.percentile(lst_outputs,[2.5,97.5]))

# lst = []
# for i in range(100000):
#     date = skewnorm.rvs(3.5,loc=750,scale=40)
#     lst.append(date)
# plt.hist(lst,bins=50,color='g',edgecolor='black',alpha=0.7)
# plt.title('Pre-WC CFO')
# plt.ylabel('Frequency',fontsize=8)
# plt.tight_layout()
# plt.show()
# print(np.mean(lst),np.std(lst))
# print(np.percentile(lst,[0.5,99.5]))
