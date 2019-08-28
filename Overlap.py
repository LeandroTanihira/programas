# -*- coding: utf-8 -*-
"""
Created on Tue Aug 27 00:10:04 2019

@author: Leand
"""

import pandas as pd
from copy import deepcopy
import time
import xlrd

inicio = time.time()

class data:

    def __init__(self):
        self.lista = []
        
    '''A função overlap vai realizar todos tratamentos necessários para tirar todos overlaps'''
    def overlap(self, linha):
            
        self.lista.append(linha)
        
        indice = len(self.lista)
        
        cepi, cepf, regiao, prazo = linha[0], linha[1], linha[4], linha[5]
        
        #É preciso ter mais de uma linha na lista
        if indice > 1:
            
            for i in range(0,len(self.lista)):
                #Linha atual está dentro da faixa anterior
                if cepi <= self.lista[i][1]:
                    if cepi >= self.lista[i][0] and cepf <= self.lista[i][1] and regiao == self.lista[i][4] and prazo == self.lista[i][5] and linha != self.lista[i]:
                        self.lista.pop(indice-1)
                        break
                    #Linha atual tem intersecção com a faixa de cep anterior
                    elif cepi <= self.lista[i][1] and cepf >= self.lista[i][1] and regiao == self.lista[i][4] and prazo == self.lista[i][5] and linha != self.lista[i]:
                        self.lista[i][1] = cepf
                        self.lista.pop(indice-1)
                        break
                    #Linha atual tem intersecção com a faixa de CEP anterior, mas região ou prazo são diferentes
                    elif cepi <= self.lista[i][1] and cepf >= self.lista[i][1] and linha != self.lista[i]:
                        if self.lista[i][0] == self.lista[i][1]:
                            self.lista[indice-1][0] = cepi + 1
                            break
                        else:
                            self.lista[indice-1][1] = cepi - 1
                            break
                    #Linha atual está dentro da faixa de CEP anterior, mas região ou prazo são deiferentes
                    elif cepf < self.lista[i][1] and linha != self.lista[i]:
                        self.lista.append(self.lista[i].copy())
                        self.lista[i][1]   = cepi - 1
                        self.lista[len(self.lista)-1][0] = cepf + 1
                        break
                    #Linha atual possui CEPI maior que o CEPF da faixa anterior
                    else:
                        pass
        #Caso só possua uma linha na lista
        else: 
            pass

file  = 'C:\\Users\\Leandro Mateus\\Desktop\\Projetos\\Python\\Overlap\\programas\\Entrada Python Overlap - Azul.xlsx'
lista = []
wb = xlrd.open_workbook(file)
sheet = wb.sheet_by_index(0)
for row_num in range(sheet.nrows):
    row_value = sheet.row_values(row_num)
    lista.append(row_value)

lista.pop(0)

listaFinal = data()

for linha in lista:
    listaFinal.overlap(linha)

saida = deepcopy(listaFinal.lista)

for i in range(len(saida)-1,-1,-1):
    if saida[i][0] > saida[i][1]:
        saida.pop(i)
    else:
        pass
    
tempo = time.time()

saida.insert(0, ['destination_zip_code_start','destination_zip_code_end','destination_state','destination_city','destination_geographic_identifier','delivery_time','minimum_total_freight','gris_min','gris_max','gris_fixed','gris_type','gris_value','gris_base','gris_fraction','gris_range_start','gris_range_end','gris_range_base','gris_inrange_fixed','gris_inrange_type','gris_inrange_value','gris_inrange_base','gris_inrange_fraction','gris_sum','gris_range_base_calculation_mode','tas_min','tas_max','tas_fixed','tas_type','tas_value','tas_base','tas_fraction','tas_range_start','tas_range_end','tas_range_base','tas_inrange_fixed','tas_inrange_type','tas_inrange_value','tas_inrange_base','tas_inrange_fraction','tas_sum','tas_range_base_calculation_mode','trt_min','trt_max','trt_fixed','trt_type','trt_value','trt_base','trt_fraction','trt_range_start','trt_range_end','trt_range_base','trt_inrange_fixed','trt_inrange_type','trt_inrange_value','trt_inrange_base','trt_inrange_fraction','trt_sum','trt_range_base_calculation_mode','tde_min','tde_max','tde_fixed','tde_type','tde_value','tde_base','tde_fraction','tde_range_start','tde_range_end','tde_range_base','tde_inrange_fixed','tde_inrange_type','tde_inrange_value','tde_inrange_base','tde_inrange_fraction','tde_sum','tde_range_base_calculation_mode','tda_min','tda_max','tda_fixed','tda_type','tda_value','tda_base','tda_fraction','tda_range_start','tda_range_end','tda_range_base','tda_inrange_fixed','tda_inrange_type','tda_inrange_value','tda_inrange_base','tda_inrange_fraction','tda_sum','tda_range_base_calculation_mode','tsb_min','tsb_max','tsb_fixed','tsb_type','tsb_value','tsb_base','tsb_fraction','tsb_range_start','tsb_range_end','tsb_range_base','tsb_inrange_fixed','tsb_inrange_type','tsb_inrange_value','tsb_inrange_base','tsb_inrange_fraction','tsb_sum','tsb_range_base_calculation_mode','suframa_min','suframa_max','suframa_fixed','suframa_type','suframa_value','suframa_base','suframa_fraction','suframa_range_start','suframa_range_end','suframa_range_base','suframa_inrange_fixed','suframa_inrange_type','suframa_inrange_value','suframa_inrange_base','suframa_inrange_fraction','suframa_sum','suframa_range_base_calculation_mode','fluvial_insurance_min','fluvial_insurance_max','fluvial_insurance_fixed','fluvial_insurance_type','fluvial_insurance_value','fluvial_insurance_base','fluvial_insurance_fraction','fluvial_insurance_range_start','fluvial_insurance_range_end','fluvial_insurance_range_base','fluvial_insurance_inrange_fixed','fluvial_insurance_inrange_type','fluvial_insurance_inrange_value','fluvial_insurance_inrange_base','fluvial_insurance_inrange_fraction','fluvial_insurance_sum','fluvial_insurance_range_base_calculation_mode','toll_min','toll_max','toll_fixed','toll_type','toll_value','toll_base','toll_fraction','toll_range_start','toll_range_end','toll_range_base','toll_inrange_fixed','toll_inrange_type','toll_inrange_value','toll_inrange_base','toll_inrange_fraction','toll_sum','toll_range_base_calculation_mode','pickup_min','pickup_max','pickup_fixed','pickup_type','pickup_value','pickup_base','pickup_fraction','pickup_range_start','pickup_range_end','pickup_range_base','pickup_inrange_fixed','pickup_inrange_type','pickup_inrange_value','pickup_inrange_base','pickup_inrange_fraction','pickup_sum','pickup_range_base_calculation_mode','delivery_min','delivery_max','delivery_fixed','delivery_type','delivery_value','delivery_base','delivery_fraction','delivery_range_start','delivery_range_end','delivery_range_base','delivery_inrange_fixed','delivery_inrange_type','delivery_inrange_value','delivery_inrange_base','delivery_inrange_fraction','delivery_sum','delivery_range_base_calculation_mode','cte_min','cte_max','cte_fixed','cte_type','cte_value','cte_base','cte_fraction','cte_range_start','cte_range_end','cte_range_base','cte_inrange_fixed','cte_inrange_type','cte_inrange_value','cte_inrange_base','cte_inrange_fraction','cte_sum','cte_range_base_calculation_mode','seccat_min','seccat_max','seccat_fixed','seccat_type','seccat_value','seccat_base','seccat_fraction','seccat_range_start','seccat_range_end','seccat_range_base','seccat_inrange_fixed','seccat_inrange_type','seccat_inrange_value','seccat_inrange_base','seccat_inrange_fraction','seccat_sum','seccat_range_base_calculation_mode','itr_min','itr_max','itr_fixed','itr_type','itr_value','itr_base','itr_fraction','itr_range_start','itr_range_end','itr_range_base','itr_inrange_fixed','itr_inrange_type','itr_inrange_value','itr_inrange_base','itr_inrange_fraction','itr_sum','itr_range_base_calculation_mode','insurance_min','insurance_max','insurance_fixed','insurance_type','insurance_value','insurance_base','insurance_fraction','insurance_range_start','insurance_range_end','insurance_range_base','insurance_inrange_fixed','insurance_inrange_type','insurance_inrange_value','insurance_inrange_base','insurance_inrange_fraction','insurance_sum','insurance_range_base_calculation_mode','ademe_min','ademe_max','ademe_fixed','ademe_type','ademe_value','ademe_base','ademe_fraction','ademe_range_start','ademe_range_end','ademe_range_base','ademe_inrange_fixed','ademe_inrange_type','ademe_inrange_value','ademe_inrange_base','ademe_inrange_fraction','ademe_sum','ademe_range_base_calculation_mode','schedule_delivery_min','schedule_delivery_max','schedule_delivery_fixed','schedule_delivery_type','schedule_delivery_value','schedule_delivery_base','schedule_delivery_fraction','schedule_delivery_range_start','schedule_delivery_range_end','schedule_delivery_range_base','schedule_delivery_inrange_fixed','schedule_delivery_inrange_type','schedule_delivery_inrange_value','schedule_delivery_inrange_base','schedule_delivery_inrange_fraction','schedule_delivery_sum','schedule_delivery_range_base_calculation_mode','reshipping_min','reshipping_max','reshipping_fixed','reshipping_type','reshipping_value','reshipping_base','reshipping_fraction','reshipping_range_start','reshipping_range_end','reshipping_range_base','reshipping_inrange_fixed','reshipping_inrange_type','reshipping_inrange_value','reshipping_inrange_base','reshipping_inrange_fraction','reshipping_sum','reshipping_range_base_calculation_mode','return_fee_min','return_fee_max','return_fee_fixed','return_fee_type','return_fee_value','return_fee_base','return_fee_fraction','return_fee_range_start','return_fee_range_end','return_fee_range_base','return_fee_inrange_fixed','return_fee_inrange_type','return_fee_inrange_value','return_fee_inrange_base','return_fee_inrange_fraction','return_fee_sum','return_fee_range_base_calculation_mode','other_fee_min','other_fee_max','other_fee_fixed','other_fee_type','other_fee_value','other_fee_base','other_fee_fraction','other_fee_range_start','other_fee_range_end','other_fee_range_base','other_fee_inrange_fixed','other_fee_inrange_type','other_fee_inrange_value','other_fee_inrange_base','other_fee_inrange_fraction','other_fee_sum','other_fee_range_base_calculation_mode'])

final = pd.DataFrame(saida)
final = final.rename(columns=final.iloc[0])
final = final.drop(0)

tempo_df = time.time()

fim = time.time()

print('Tempo de execução', fim - inicio)
    