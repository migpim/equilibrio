# equilibrio
Extrator de dados da consulta de equilíbrio foi criado com o objetivo de extrair informação de um ficheiro .xlsx automaticamente gerado quando um profissional de saúde preenche um forms no Microsoft Forms da sua instituição. O programa gera um ficheiro .docx e copia a informação clínica para o clipboard de modo a que o clínico apenas necessite de colar no software utilizado pela sua instituição a informação gerada por si durante a observação do doente.

- O extractor de dados está programado para pesquisar um ficheiro excel com um nome com o formato 'Consulta de Equilíbrio(*).xlsx', tal como é gerado pelo Microsoft forms.
- O programa guarda a informação num ficheiro word com o nome output_(numero do processo).docx com base no template já previamente definido em template.docx

Para que o programa corra sem erros é necessário que no mesmo diretório estejam guardados o programa extractor_py.py, um (e apenas um!) ficheiro excel com o nome no formato 'Consulta de Equilíbrio(*).xlsx' e o ficheiro template.docx
