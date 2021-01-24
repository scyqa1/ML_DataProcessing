import numpy as np
from numpy import transpose
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.impute import SimpleImputer  # 填充缺失值的类
from sklearn.ensemble import RandomForestRegressor  # 随机森林回归
from sklearn.linear_model import BayesianRidge
from sklearn.model_selection import cross_val_score  # 交叉验证
import xlrd
import xlwt
from sklearn.neighbors import KNeighborsRegressor
from sklearn.pipeline import make_pipeline
from sklearn.tree import DecisionTreeRegressor

io = r'Lake_v2.xls'
f = pd.read_excel(io, sheet_name=0, usecols=[0, 1, 4])
#f.head()
data = f[['Year', 'Month','TotalP']]

known_target = data[data.TotalP.notnull()].values
unknown_target = data[data.TotalP.isnull()].values

y=known_target[:,2]#y是年龄，第一列数据
x=known_target[:,[0,1]]#x是特征属性值，后面几列

rfr=RandomForestRegressor(random_state=0,n_estimators=2000,n_jobs=-1)

#根据已有数据去拟合随机森林模型
rfr.fit(x,y)
#预测缺失值
predicted = rfr.predict(unknown_target[:,[0,1]])
#填补缺失值
f.loc[(f.TotalP.isnull()),'TotalP'] = predicted
#把数值型特征都放到随机森林里面去

print(predicted)
print(f)

#写入文件
fw = xlwt.Workbook()
sheetw = fw.add_sheet('randomforest', cell_overwrite_ok=True)

sheetw.write(0, 0, 'Year')  # 第1行第1列
sheetw.write(0, 1, 'Month')  # 第2行第1列
sheetw.write(0, 2, 'TotalP')

for i in range(len(f)):
    d = f.loc[i]
    sheetw.write(i + 1, 0, d[0])  # 第1行第1列
    sheetw.write(i + 1, 1, d[1])  # 第2行第1列
    sheetw.write(i + 1, 2, d[2])  # 第3行第1列
fw.save('Lake_v3.xls')

io = r'Lake_v3.xls'
f = pd.read_excel(io, sheet_name=0, usecols=[0, 1, 2])
f.head()
data = f[['Year', 'Month','TotalP']]

#随机森林
X_full, y_full = np.round(data.values[:,[0,1]],6), np.round(data.values[:,2],6)


n_samples = X_full.shape[0]
n_features = X_full.shape[1]


rng = np.random.RandomState(0)
missing_rate = 0.3
n_missing_samples = int(np.floor(n_samples * n_features * missing_rate))

missing_features = rng.randint(0, n_features, n_missing_samples)
missing_samples = rng.randint(0, n_samples, n_missing_samples)

# missing_samples = rng.choice(dataset.data.shape[0],n_missing_samples,replace=False)

X_missing = X_full.copy()
y_missing = y_full.copy()

X_missing[missing_samples, missing_features] = np.nan

X_missing[missing_samples, missing_features] = np.nan

X_missing = pd.DataFrame(X_missing)

# 使用均值进行填补
from sklearn.impute import SimpleImputer

imp_mean = SimpleImputer(missing_values=np.nan, strategy='mean')
X_missing_mean = np.round(imp_mean.fit_transform(X_missing),6)

# 查看缺失值填补之后每列的缺失值个数
pd.DataFrame(X_missing_mean).isnull().sum()



################使用随机森林回归填补缺失值#################

X_missing_reg = X_missing.copy()

sortindex = np.argsort(X_missing_reg.isnull().sum(axis=0)).values

for i in sortindex:
    # 构建我们的新特征矩阵和新标签
    df = X_missing_reg
    fillc = df.iloc[:, i]
    df = pd.concat([df.iloc[:, df.columns != i], pd.DataFrame(y_full)], axis=1)
    # 在新特征矩阵中，对含有缺失值的列，进行0的填补
    df_0 = SimpleImputer(missing_values=np.nan,
                         strategy='constant', fill_value=0).fit_transform(df)
    # 找出我们的训练集和测试集
    Ytrain = fillc[fillc.notnull()]
    Ytest = fillc[fillc.isnull()]
    Xtrain = df_0[Ytrain.index, :]
    Xtest = df_0[Ytest.index, :]
    # 用随机森林回归来填补缺失值
    rfc = RandomForestRegressor(n_estimators=100)
    rfc = rfc.fit(Xtrain, Ytrain)
    Ypredict = np.round(rfc.predict(Xtest),6)
    # 将填补好的特征返回到我们的原始的特征矩阵中
    X_missing_reg.loc[X_missing_reg.iloc[:, i].isnull(), i] = Ypredict

# 所有列的缺失值填补完整

X_missing_reg.isnull().sum()

#mse计算
X = [X_full,X_missing_mean,X_missing_reg]

mse = []
for i in X:
    estimator = RandomForestRegressor(random_state=0, n_estimators=1000)
    scores = cross_val_score(estimator,i,y_full,scoring='neg_mean_squared_error',cv=5).mean()
    mse.append(scores * -1)

print(mse)

#画图
x_labels = ['Full data',
            'Mean Imputation',
            'Regressor Imputation']

colors = ['r', 'g', 'b', 'orange']
plt.figure(figsize=(12, 6))
ax = plt.subplot(111) #添加子图
for i in np.arange(len(mse)):
    ax.barh(i, mse[i],color=colors[i], alpha=0.6, align='center')

ax.set_title('Imputation Techniques with China Lake')
ax.set_xlim(left=np.min(mse) * 0.9,right=np.max(mse) * 1.1)
ax.set_yticks(np.arange(len(mse)))
ax.set_xlabel('MSE')
ax.invert_yaxis()
ax.set_yticklabels(x_labels)
plt.show()