import pandas as pd

def read_excel_data(input_read):
    try:
        df = pd.read_excel(input_read)
        return df
    except FileNotFoundError:
        print(f"Arquivo '{input_read}' não encontrado.")
        return None
    except Exception as e:
        print(f"Erro ao ler o arquivo '{input_read}': {e}")
        return None

def save_to_txt(df, input_data, output_data, filtered_name):
    try:
        lower_filtered_name = filtered_name.lower().strip()
        df_filtered = df[df['nome'].str.lower().str.strip() == lower_filtered_name]
        if not df_filtered.empty:
            desired_columns = [column for column in df.columns if column != 'status']
            formatted_data = df_filtered[desired_columns].to_string(index=False, float_format='%.0f')
        else:
            formatted_data = f"Não há dados para o nome '{filtered_name}'."
        with open(output_data, 'w', encoding='utf-8') as txt_file:
            txt_file.write(formatted_data)
        df.loc[df['nome'].str.lower().str.strip() == lower_filtered_name, 'status'] = 'OK'
        df.to_excel(input_data, index=False)
    except Exception as e:
        print(f"Erro ao salvar os dados: {e}")
 
def is_empty_string(value):
    return value.strip() == ""
    
def create_record(df, name, age, job, sex):
    try:
        while is_empty_string(name) or not isinstance(age, int) or is_empty_string(job) or is_empty_string(sex):
            print("Todos os campos são obrigatórios. Por favor, insira valores válidos.")
            name = input("Digite o nome: ")
            age = input("Digite a idade: ")
            job = input("Digite a profissão: ")
            sex = input("Digite o sexo: ")
            try:
                age = int(age)
            except ValueError:
                age = None

        if 'id' not in df.columns:
            df['id'] = range(1, len(df) + 1) if not df.empty else [1]
        max_id = df['id'].max() if not df.empty else 0
        new_id = max_id + 1
        new_record = {'id': new_id, 'nome': name, 'idade': age, 'profissao': job, 'sexo': sex, 'status': 'Created'}
        df = pd.concat([df, pd.DataFrame([new_record])], ignore_index=True)
        return df
    except Exception as e:
        print(f"Erro ao criar o registro: {e}")
        return None

def get_name_by_id(df, id):
    try:
        name = df.loc[df['id'] == id, 'nome'].iloc[0]
        return name
    except IndexError:
        return None

def update_record_by_id(df, id, new_name):
    try:
        if id not in df['id'].values:
            print(f"Erro ao atualizar o registro: O ID {id} não existe.")
            return df
        df.loc[df['id'] == id, 'nome'] = new_name
        return df
    except Exception as e:
        print(f"Erro ao atualizar o registro: {e}")
        return None

def delete_record(df, id):
    try:
        if id not in df['id'].values:
            print(f"Erro ao excluir o registro com o ID {id}!")
            return df
        df = df[df['id'] != id]
        return df
    except Exception as e:
        print(f"Erro ao excluir o registro: {e}")
        return None
