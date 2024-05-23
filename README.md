# Zer0Mail

Zer0Mail é uma automação para gerar relatórios de toners e enviá-los por e-mail para seu chefe. Além disso, agora possui um menu interativo para facilitar a utilização e novas funcionalidades como adição de impressoras, relacionamento entre impressoras e toners, e atualização da tabela.

## Requisitos

- Python 3.6+
- Pandas
- smtplib
- openpyxl

## Instalação

1. Clone este repositório:
    ```sh
    git clone https://github.com/usuario/Zer0Mail.git
    cd Zer0Mail
    ```

2. Crie um ambiente virtual e instale as dependências:
    ```sh
    python -m venv venv
    source venv/bin/activate  # No Windows, use `venv\Scripts\activate`
    pip install pandas openpyxl
    ```

## Uso

Execute o aplicativo:
```sh
python app.py
```
Você verá o menu principal com as seguintes opções:
```bash
Menu Principal
1. Adicionar Impressora e Toner
2. Atualizar tabela
3. Quantificar toners
4. Relacionar Toners a Impressoras
5. Enviar tabela para email
6. Sair
```

## Opções do Menu
### Adicionar Impressora e Toner

1. Sub-opções:
    1. Adicionar Impressora
    2. Adicionar Toner
    3. Adiciona uma nova impressora ou toner ao sistema, salvando-os em arquivos de texto.

2. Atualizar tabela

    Atualiza a tabela de toners com quantidade, destino e impressora.

3. Quantificar toners

    Exibe a quantidade total de cada toner.

4. Relacionar Toners a Impressoras

    Exibe os toners relacionados a uma impressora específica.

5. Enviar tabela para email

    Converte a tabela de CSV para Excel e envia o relatório como anexo por e-mail.

6. Sair

    Sai do aplicativo.

## Estrutura do Código
### Funções Principais
```python
# Configura e envia um e-mail com o relatório anexado.
send_email(from_email, to_email, subject, body, attachment_filename)

# Ajusta automaticamente a largura das colunas de um arquivo Excel.
auto_adjust_column_width(book)

# Converte um arquivo CSV em um arquivo Excel e ajusta a largura das colunas.
csv_to_excel(csv_filename, excel_filename)

# Adiciona uma nova impressora ou toner ao sistema.
add_impressora_toner()

# Atualiza a tabela de toners com quantidade, destino e impressora.
update_table()

# Exibe a quantidade total de cada toner.
quantificar_toners()

# Exibe os toners relacionados a uma impressora específica.
relacionar_toners_impressoras()

# Converte o CSV em Excel e envia o arquivo Excel como anexo para um endereço de e-mail especificado.
enviar_tabela_email()

# Exibe o menu principal e chama as funções apropriadas com base na escolha do usuário.
main_menu()

```

# Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir um problema ou enviar um pull request.

# Licença
Este projeto está licenciado sob a [Licença MIT](https://github.com/Zer0G0ld/Zer0Mail/blob/main/LICENSE)
