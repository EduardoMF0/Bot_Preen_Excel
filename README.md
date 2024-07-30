## Bot_Preen_Excel
### Bot para verificação do sistema por meio de imagens e preenchimento de uma planilha com dados de mais de 5 mil funcionários.

### Tarefas:

O bot identifica elementos por imagem e está conectado a uma planilha Excel. Dentro do sistema, ele pesquisa o colaborador pelo ID e verifica se o nome no perfil está correto. Em seguida, clica no ícone que abre a página de documentos e verifica se o documento necessário está presente. Se o documento existir, o bot faz o download e preenche a planilha indicando a presença ou ausência do documento.

### Linguagem:

Python. 

**Bibliotecas**: Pyautogui, openpyxl, Time, Tkinter, pyperclip.

### Motivo:

Projeto desenvolvido para a antiga empresa onde trabalhava. Nosso cliente solicitou com urgência um documento específico de cerca de cinco mil colaboradores de seu produto. Esse documento estava no perfil de cada funcionário. A tarefa consistia em pegar o ID do colaborador na planilha, acessar seu perfil, verificar cada um dos seus documentos, identificar a presença ou ausência do documento específico e, dependendo do resultado, preencher a planilha correspondente.

Para resolver essa demanda, meus superiores alocaram aproximadamente sete pessoas do nosso setor administrativo. No entanto, essa equipe levaria mais de quatro semanas para concluir a tarefa, impossibilitando a entrega dentro do prazo acordado. Além disso, os encarregados não estariam disponíveis todos os dias na empresa e não havia possibilidade de trabalho remoto.

Diante disso, tive a ideia de criar um bot para automatizar essa tarefa. Com a aprovação dos meus superiores, desenvolvi e implementei o bot, que foi testado e utilizado com sucesso para concluir a demanda dentro do prazo estabelecido.
