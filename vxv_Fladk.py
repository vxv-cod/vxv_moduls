import os
import pyodbc
from sqlalchemy import create_engine, text, MetaData, Table, Column, Integer, String
from sqlalchemy.engine import URL



def SQlServer_pyodbc_select(table, item):
    '''Сбор данных из базы данных по имени таблицы и строки списка колонок'''

    param = "Driver={SQL Server};Server=tnnc-sapsan-db;Database=SapsanPlus;"
    with pyodbc.connect(param) as connection:
        columnName = item.split(', ')
        dbCursor = connection.cursor()
        dbCursor.execute(f'SELECT {item} FROM dbo.{table}')
        data = [{columnName[i] : str(row[i]) for i in range(len(columnName))} for row in dbCursor.fetchall()]
    return data



def SQlServer_sqlalchemy_select(table, item):
    '''Сбор данных из базы данных по имени таблицы и строки списка колонок'''
    columnName = item.split(', ')
    param = "Driver={SQL Server};Server=tnnc-sapsan-db;Database=SapsanPlus;"
    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": param})
    engine = create_engine(connection_url)
    with engine.connect() as conn:
        result = conn.execute(text(f'''SELECT {item} FROM dbo.{table}'''))
        data = [{columnName[i] : str(row[i]) for i in range(len(columnName))} for row in result.all()]
    return data



def table_objict(columns_name, table_name, metadata):
    '''Описание таблицы как объекта'''
    param = [table_name, metadata]
    for i, v in enumerate(columns_name):
        if i == 0:
            param.append(Column(columns_name[i], Integer, primary_key=True))
        else:
            param.append(Column(columns_name[i], String))
    user_table = Table(*param)
    return user_table



def create_db_tab(basedir, db_name, table_name, columns_name):
    '''Создание таблицы в базе данных'''    
    # engine = create_engine("sqlite+pysqlite:///:memory:", echo=True, future=True)
    engine = create_engine("sqlite:///" + os.path.join(basedir, f'{db_name}.db'), echo=True, future=True)
    metadata = MetaData()
    table_objict(columns_name, table_name, metadata)
    metadata.create_all(engine)



def insert_db_tab(basedir, db_name, table_name, columns_name, newdata):
    '''Вставка данных в таблицу БД'''
    engine = create_engine("sqlite:///" + os.path.join(basedir, f'{db_name}.db'), echo=True, future=True)
    metadata = MetaData()
    user_table = table_objict(columns_name, table_name, metadata)
    
    kwargs = {user_table.c.keys()[i] : v for i, v in enumerate(newdata, 1)}

    with engine.connect() as conn:
        req = user_table.insert().values(**kwargs)
        print('req----- ', req)
        conn.execute(req)
        conn.commit()



def select_db_tab(basedir, db_name, table_name, columns_name):
    '''Извлечение данных из таблицы БД'''
    engine = create_engine("sqlite:///" + os.path.join(basedir, f'{db_name}.db'), echo=True, future=True)
    metadata = MetaData()
    user_table = table_objict(columns_name, table_name, metadata)

    with engine.connect() as conn:
        req = user_table.select()
        print(req.compile())
        data = conn.execute(req)
        # data = {row for row in data.fetchall()}
        data = [{columns_name[i] : str(row[i]) for i in range(len(columns_name))} for row in data.fetchall()]

        print(data)




if __name__ == "__main__":
    # table = 'contract'
    # item ='id, shifr, subject'
    # print(SQlServer_sqlalchemy_select(table, item))
    # print(SQlServer_pyodbc_select(table, item))

    # create_db_tab(
    #     basedir=r'C:\vxvproj\vxv_Flask\WebDen', 
    #     db_name='main111',
    #     # table_name='user_table',
    #     table_name='user_table3',
    #     columns_name=['id', 'name', 'fullname']
    #     )


    # insert_db_tab(
    #     basedir=r'C:\vxvproj\vxv_Flask\WebDen', 
    #     db_name='main111',
    #     table_name='user_table222',
    #     columns_name=['id', 'name', 'fullname'],
    #     newdata = ['AAAAA', 'HHHHHHHHH']
    #     )

    select_db_tab(
        basedir=r'C:\vxvproj\vxv_Flask\WebDen', 
        db_name='main111',
        # table_name='user_table',
        table_name='user_table222',
        columns_name=['id', 'name', 'fullname']
        )

    pass
