@echo off
echo Executando Exportacao…

echo mysqldump.exe  --lock-tables --hex-blob --opt --flush-logs --triggers --port=3306  %1 cidade conjuge corretor empresa igpm imovel incorporador incorporador_loteamento loteamento memorial_descritivo part part2 participante plano_contas price procuradores profissao quadras representantes tabela tipodocumento vendedor --character-set=utf8 --user=%2 -h %3 --password=%4 --opt -v> "%5"
mysqldump.exe  --lock-tables --hex-blob --opt --flush-logs --triggers --port=3306  %1  cidade conta_bancaria conjuge corretor empresa igpm imovel incorporador incorporador_loteamento loteamento memorial_descritivo part part2  participante plano_contas price procuradores profissao quadras representantes representa_participante tabela tipodocumento vendedor  --character-set=utf8 --user=%2 -h %3 --password=%4 --opt -v> "%5"

echo Feito! 
