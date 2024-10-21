**Descrição do Projeto: Exportador de Dados em Delphi**

Este projeto foi desenvolvido em Delphi e tem como objetivo permitir a exportação de dados de um banco de dados Firebird para arquivos nos formatos `.xls` (Excel) e `.csv` (valores separados por vírgulas).

### Funcionalidades:
1. **Conexão com o Banco de Dados Firebird**:  
   O sistema se conecta a um banco de dados Firebird, permitindo acesso a todas as tabelas contidas no banco.
   
2. **Listagem de Tabelas**:  
   Após a conexão, o aplicativo exibe uma lista de todas as tabelas disponíveis no banco de dados.
   
3. **Seleção de Tabelas**:  
   O usuário pode marcar ou desmarcar as tabelas desejadas para a exportação. A interface permite a seleção individual ou múltipla de tabelas.
   
4. **Exportação de Registros**:  
   Os registros das tabelas selecionadas são exportados em dois formatos:
   - **.xls**: Para visualização e manipulação em planilhas do Excel.
   - **.csv**: Para uso em outros sistemas ou manipulação de dados de forma simplificada.
   
5. **Interface Intuitiva**:  
   A interface foi projetada para ser simples e intuitiva, facilitando a navegação e a seleção das tabelas para exportação.

### Tecnologias Utilizadas:
- **Delphi**: Linguagem de programação para o desenvolvimento da aplicação.
- **Firebird**: Banco de dados relacional utilizado para armazenar os dados.
- **Firedac**: Componente para a conexão com o banco de dados Firebird.
- **Excel e CSV**: Código customizado para gerar os arquivos `.xls` e `.csv`.

Este projeto oferece uma solução simples e eficiente para a exportação de dados de um banco de dados Firebird, sendo ideal para administradores de dados e desenvolvedores que precisem realizar backups ou migrações de dados de maneira rápida e personalizada.
