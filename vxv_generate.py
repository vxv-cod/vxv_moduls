from sqlacodegen_v2.external import generate_models


def generate_model_fun(outfile = None):
    from sqlalchemy import create_engine
    from sqlalchemy.orm import declarative_base
    from sqlacodegen_v2.generators import DeclarativeGenerator, SQLModelGenerator
    
    engine = create_engine("sqlite:///sqlite.db")
    metadata = declarative_base().metadata
    metadata.reflect(bind=engine)
    generator = DeclarativeGenerator(metadata, bind=engine, options=[])
    with open(outfile, 'w') as f:
        f.write(generator.generate())  


def generate_model_all():
    generators = ['dataclasses', 'declarative', 'declarative-dataclasses', 'sqlmodels', 'tables']
    db_url = "sqlite:///sqlite.db"

    for generator in generators:
        outfile_path = f"models_{generator}.py"
        generate_models(db_url=db_url, generator=generator, outfile_path=outfile_path)


def generate_model(db_url:str, generator:str):
    '''generators = ['dataclasses', 'declarative', 'declarative-dataclasses', 'sqlmodels', 'tables']'''
    generator = 'declarative'
    outfile_path = f"db_models_{generator}.py"
    generate_models(db_url, generator, outfile_path=outfile_path)


if __name__ == '__main__':
    # generate_model_fun('db_generate_model_fun.py')
    # generate_model_all()
    # db_url = "sqlite:///sqlite.db"
    generate_model(db_url, 'declarative')



'''Для Алхимии версии 1.4'''
# import io
# import sys
# from sqlalchemy import create_engine, MetaData
# from sqlacodegen.codegen import CodeGenerator

# def generate_model(outfile = None):

#     engine = create_engine(f'sqlite:///sqlite.db')
#     metadata = MetaData(bind=engine)
#     metadata.reflect()
#     outfile = io.open(outfile, 'w', encoding='utf-8') if outfile else sys.stdout
#     generator = CodeGenerator(metadata)
#     generator.render(outfile)

# if __name__ == '__main__':
#     generate_model('db1.py')