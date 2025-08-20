# Inadimplência
Me foi demandado um painel  para gestão de clientes inadimplentes da empresa, utilizando uma conexão com a ferramenta Conta Azul, usada pela empresa para gestão monetária interna, para assim, a equipe de atendimento ter uma visão mais direta e básica aos clientes não pagantes para cobranças e medidas necessárias, com algumas regras de negócio:

- Arquivo que consiga alimentar manualmente de acordo com as conversas com clientes, juntando uma base automática a uma alimentada manualmente.

Nesse projeto utilizei as seguintes ferramentas e técnicas:

- Python (com as bibliotecas, Pandas, Requests, OS, Json);
- Python + Selenium (para RPA);
- Excel online;
- Macros dentro do Excel;
- PowerBI;
- Uso da API do Conta Azul;
- Medidas e colunas DAX, Tooltips no PowerBI.

Para ser possível atender a todas regras de negócio precisei dividir em algumas partes:

1º -> Script utilizando a API do Conta Azul para conseguir resgatar os dados da aba em específico de inadimplentes, subindo todos essas dados para uma planilha base com todos os dados

2º -> Criação de uma planilha dimensão, que puxa os nomes da planilha base e criando uma tabela dinâmica em outra aba, para com novos nomes a tabela dinâmica seja atualizada automaticamente e junto a ela com os nomes com essa atualização colunas com dados que serão preenchidos manualmente, com validação de dados para evitar erros de digitação, para contornar problemas, como, clientes que deixam de ser inadimplentes e depois voltam a ser, novos clientes, e outras situações como essas para não quebrar a relação das linhas entre o nome que vem de forma automática da API e essas informações que serão alimentadas manualmente eu segui as seguintes métricas:

Divisão entre tabelas fato (base dos dados que vem da API) e dimensão (essa que pega esses dados base e juntamente alimentada por informações manualmente);

Dentro da dimensão para evitar erros de consistência e correlação entre as informações, separei em 2 abas, a primeira aba "Verificação" puxa os nomes que estão na tabela base e contém 2 macros, um para chamar o Script que faz a chamada a API do Conta Azul e faz toda atualização da base de dados e outro macro que faz uma verificação, com os nomes novos, ausentes e que voltaram, fazendo coisas como, nomes que foram ausentes e voltaram, busca o nome no log para conseguir tirar essa verificação e se sim deixa em branco as informações que foram colocadas manualmente, quando o nome é ausente as informações que foram colocadas manualmente tornam a ser "APAGADAS!" e o nome não é apagado da tabela se não quebraria a consistência ele somente fica pintado de vermelho, na aba "Consulta" é onde se encontra minha tabela dinâmica que vem da aba "Verificação" e as colunas que são colocadas informações manualmente, com verificação de dados para evitar dados diferentes do esperado

3º -> RPA para atualizar o dataset desse painel e assim não precisando assinar nenhum plano do PowerBI
