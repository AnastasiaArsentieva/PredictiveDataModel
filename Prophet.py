import pandas as pd
import numpy as np
import openpyxl
import xlsxwriter
import matplotlib.pyplot as plt
from prophet import Prophet


def prophet(file_exal):
    #цикл считывания и обработки всех столбцов файла
    #df_all = pd.read_excel(file_exal)
    df_all = file_exal
    name_shares = df_all.columns
    name_shares = name_shares.drop(['TRADEDATE'])
    #print(name_shares)

    #функция для вычисления стандартной метрики SMAPE
    def standard_smape(actual,forecast):
        return round((np.mean(np.abs(forecast - actual) / (np.abs(actual) + np.abs(forecast))))*100,1)

    #количество прогнозных значений
    HORIZONT = 32
    #номер столбца
    i = 1

    for i in name_shares:
        #загружаем данные
        df = pd.DataFrame(df_all, columns=['TRADEDATE',i])
        df['TRADEDATE'] = pd.to_datetime(df['TRADEDATE'])
        df.columns = ['ds','y']

    #визуализируем ряд (замутила)
    #plt.figure(figsize=(10,6))
    #plt.scatter(pd.to_datetime(df['ds']), df['y'],s=1,c='#0072B2')
    #plt.xlabel('Дата')
    #plt.ylabel(i)
    #plt.show()

        #создаём модель Prophet
        model = Prophet()
        #обучаем модель
        model.fit(df)

        future = model.make_future_dataframe(periods=32)
        #print(future)

        #получаем прогнозы
        forecast = model.predict(future)
        itog = pd.DataFrame(forecast, columns=['ds','yhat'])
        print(forecast)
        print(list(forecast))
        print(itog)

    #смотрим компоненты прогнозов (замутила)
    #fig2 = model.plot_components(forecast)
    #plt.show()

        smape = standard_smape(df['y'],itog['yhat'][:-32])
        print(f'SMAPE по {i}: {smape:.3f}')

        #Запись значений в файл
        #boards_sheets = {'все_прогнозы': forecast, 'дата_цена': itog}

#        writer = pd.ExcelWriter(i+'_прогноз.xlsx', engine='xlsxwriter')

 #       for sheet_name in boards_sheets.keys():
  #          boards_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

   #     writer.close()

    return forecast, itog