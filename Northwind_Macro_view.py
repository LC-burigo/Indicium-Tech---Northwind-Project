import pandas
import pymysql
import openpyxl

Connection = Connection = pymysql.connect(host='localhost', user='root', password='9479854532441919Lb', db='northwind')
Cursor = Connection.cursor()


def datas_ordens ():
    global Connection
    global Cursor

    Cursor.execute("SELECT DISTINCT order_date FROM order_details join orders on orders.ID = order_details.ID")
    Connection.commit()
    datas = Cursor.fetchall()
    return list(set([data[0][3:] for data in list(datas)]))

def qtd_categorias_competencia ():
    global Connection
    global Cursor

    datas = datas_ordens()
    lista_grande = []

    for data in datas:
        Cursor.execute("select products.category_id, count(*) As quantidade_de_vendas from order_details join orders on orders.ID = order_details.ID join products on order_details.product_id = products.ID where order_date like '%{}' group by products.category_id order by products.category_id".format(data))
        Connection.commit()
        qtd_ct_cmpt = Cursor.fetchall()
        lista_pequena = [element[1] for element in qtd_ct_cmpt]
        lista_pequena.insert(0, data)
        lista_grande.append(lista_pequena)
    
    return lista_grande

def macro_dataframe ():

    main_dataframe = pandas.DataFrame(qtd_categorias_competencia(), columns=[
                                  'Datas', 'Categoria 1', 'Categoria 2', 'Categoria 3', 'Categoria 4', 'Categoria 5', 'Categoria 6', 'Categoria 7', 'Categoria 8'])
    return main_dataframe
#----------------------------------------------------------------------------------------------------------------------------------------------------
def pior_categoria_competencia ():
    global Connection
    global Cursor

    lista_grande = []

    Cursor.execute("select products.category_id, products.ID, orders.shipped_date, products.unit_price, products.units_in_stock, products.units_on_order from order_details join orders on orders.ID = order_details.ID join products on order_details.product_id = products.ID where order_date like '%05/1998' GROUP BY products.ID order by products.category_id ASC, products.ID ASC")
    Connection.commit()
    pior_ct_cmpt = Cursor.fetchall()
    for element in pior_ct_cmpt:
        lista_grande.append(list(element))

    return lista_grande

def micro_dataframe ():
    derivate_dataframe = pandas.DataFrame(pior_categoria_competencia(), columns=[
        'category_id', 'product_id', 'shipped_date', 'unit_price', 'units_in_stock', 'units_on_order'])
    return derivate_dataframe
# ----------------------------------------------------------------------------------------------------------------------------------------------------
def ref_categoria_competencia ():
    global Connection
    global Cursor

    lista_grande = []

    Cursor.execute("select products.category_id, products.ID, orders.shipped_date, products.unit_price, products.units_in_stock, products.units_on_order from order_details join orders on orders.ID = order_details.ID join products on order_details.product_id = products.ID where order_date like '%04/1998' GROUP BY products.ID order by products.category_id ASC, products.ID ASC")
    Connection.commit()
    pior_ct_cmpt = Cursor.fetchall()
    for element in pior_ct_cmpt:
        lista_grande.append(list(element))

    return lista_grande

def ref_dataframe():
    derivate_dataframe = pandas.DataFrame(ref_categoria_competencia(), columns=[
        'category_id', 'product_id', 'shipped_date', 'unit_price', 'units_in_stock', 'units_on_order'])
    return derivate_dataframe
# ----------------------------------------------------------------------------------------------------------------------------------------------------
def planilha ():
    main_dataframe = macro_dataframe()
    detail_dataframe_pior = micro_dataframe()
    detail_dataframe_ref = ref_dataframe()

    with pandas.ExcelWriter('C:/Users/burig/OneDrive/Documentos/Desafio técnico Indicium/Northwind.xlsx') as writer:
        main_dataframe.to_excel(writer, sheet_name="Macro Statistics")
        detail_dataframe_pior.to_excel(writer, sheet_name="pior Statistics")
        detail_dataframe_ref.to_excel(writer, sheet_name="ref Statistics")
        
planilha()