Ajustes necessários:
linha 21 - adicionar gmail e senha SMTP
linha 54 e 56 - Inserir o gmail que irá disparar os emails e o assunto (Não mexer na linha 55)
linha 60 - Adicionar texto do corpo do email
linha 80 - Colocar o email(Mesmo da linha 54) no lugar de seuEmail@gmail


Para rodar o programa basta realizar os ajustes no código python e no terminal, dentro do diretorio rodar
pip install pyinstaller
pyinstaller --onefile certificate.py


Na pasta dist adicionar a planilha contendo os emails e, posteriormente, rodar o certificate.exe contido na pasta. 
Os arquivos baseCertificate, certificado não devem ser removidos da pasta dist
A planilha deve estar contida dentro da pasta dist
