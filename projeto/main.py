from operations import read_excel_data, create_record, get_name_by_id, update_record_by_id, delete_record, save_to_txt

if __name__ == "__main__":
    input_data = "data/input.xlsx"
    output_data = "data/output.txt"
    df = read_excel_data(input_data)

    if df is not None:
        while True:
            print("\nEscolha uma opção:")
            print("1. Criar um novo registro")
            print("2. Atualizar um registro")
            print("3. Excluir um registro")
            print("4. Filtrar por nome e salvar em TXT")
            print("5. Sair")

            choice = input("Digite o número da opção desejada: ")

            if choice == "1":
                name = input("Digite o nome: ")
                age = int(input("Digite a idade: "))
                job = input("Digite a profissão: ")
                sex = input("Digite o sexo: ")
                df = create_record(df, name, age, job, sex)
                df.to_excel(input_data, index=False)
                print(f"Registro para '{name}' criado com sucesso!")

            elif choice == "2":
                id = int(input("Digite o ID do registro a ser atualizado: "))
                new_name = input("Digite o novo nome: ")
                old_name = get_name_by_id(df, id)
                if old_name:
                    df = update_record_by_id(df, id, new_name)
                    if df is not None:
                        df.to_excel(input_data, index=False)
                        print(f"Registro com ID {id} e nome '{old_name}' atualizado para '{new_name}' com sucesso!")
                else:
                    print(f"Registro com ID {id} não encontrado.")

            elif choice == "3":
                id_to_delete = int(input("Digite o ID do registro a ser excluído: "))
                df_before_delete = df.copy()
                df = delete_record(df, id_to_delete)
                if df.equals(df_before_delete):
                    print(f"O registro com ID {id_to_delete} não existe ou não pode ser modificado.")
                else:
                    df.to_excel(input_data, index=False)
                    print(f"Registro com ID {id_to_delete} excluído com sucesso!")
                
            elif choice == "4":
                filtered_name = input("Digite o nome que deseja filtrar: ")
                save_to_txt(df, input_data, output_data, filtered_name)
                print(f"Dados filtrados para o nome '{filtered_name}' salvos em '{output_data}'.")

            elif choice == "5":
                print("Saindo...")
                break

            else:
                print("Opção inválida. Tente novamente.")
