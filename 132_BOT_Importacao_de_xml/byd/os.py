import os


class OS:
    """Métodos para manipular arquivos utilizando a biblioteca os."""

    def delete_all_files(self, folder_path):
        """Deleta todos os arquivos da pasta informada."""
        files = os.listdir(folder_path)

        for file in files:
            file_path = f"{folder_path}\\{file}"
            if os.path.isfile(file_path):
                os.remove(file_path)

    def path_exists(self, path):
        """Verifica se a pasta ou o arquivo existe e retorna True ou False."""
        return os.path.exists(path)