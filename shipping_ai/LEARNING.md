# Aprendizado de volumetria

O algoritmo e as regras de `packing.py` continuam sendo a base de validacao.
O historico e consultado depois da recomendacao normal. Cada candidato passa
por validacao antes de participar da votacao; um candidato rejeitado nao impede
que outra referencia valida seja utilizada.

Perfis fisicos usam medidas ordenadas, peso exato, familia de pack, quantidade
do pack e suas medidas. Nenhuma tolerancia foi introduzida. Os SKUs originais
permanecem no pedido e no historico. A consolidacao fisica serve para encontrar
referencias; a verificacao de capacidade continua usando os SKUs individuais,
preservando o comportamento de pack por produto.

`assignments_json` precisa cobrir exatamente o pedido. Quantidades zero geradas
pela interface sao permitidas. A distribuicao e convertida para os SKUs do novo
pedido e preenche a volumetria interativa e os modelos 3D. Planos legados com
caixas distintas, sem atribuicoes, nao sao reutilizados automaticamente.

Pedidos com itens adicionais podem aproveitar uma distribuicao manual: os
itens novos sao incluidos na ultima embalagem e o plano inteiro e revalidado.
Se nao couberem, o algoritmo atual decide. Isso nao cria uma nova otimizacao
de multi-volume. Quantidades maiores de um perfil conhecido sao rejeitadas.

A prioridade e: SKU exato com distribuicao manual; perfil fisico exato;
distribuicao manual com itens adicionais; quantidade semelhante. Dentro do
nivel escolhido, evidencias manuais pesam 2 e legadas pesam 1, multiplicadas
pela similaridade. Confianca e participacao dos votos, nao probabilidade de
encaixe. Referencias semelhantes ficam limitadas a 95%, e com itens adicionais
a 90%. O proprio pedido nao participa de sua recomendacao.

A auditoria aparece na tela do pedido e informa referencias, similaridade,
conflitos e rejeicoes. Salvar a volumetria interativa tambem verifica a cobertura
do pedido e a capacidade das embalagens.

## Verificacao

Execute na pasta `shipping_ai`:

```powershell
py -W ignore::ResourceWarning -m unittest test_learning -v
py -W ignore::ResourceWarning audit_learning.py database.db --limit 30
py learning_metrics.py database.db
```

O replay abre a base original somente para leitura e trabalha em uma copia
temporaria. Usa as regras de capacidade de `packing.py`; as restricoes de caixa
padrao de placa-mae e SSD M.2 tambem sao verificadas no fluxo da aplicacao.
O replay nao comprova uma reducao de correcoes em operacao.

## Indicadores e limites

`learning_observations` e criada quando um pedido ainda sem plano recebe uma
recomendacao. Reabrir a tela nao multiplica a contagem. A confirmacao interativa
registra se caixas e quantidades foram aceitas ou corrigidas. A comparacao dos
indicadores considera o plano de caixas, nao mudancas de posicionamento ou
redistribuicao de itens entre as mesmas caixas. A tabela guarda a observacao
por pedido; nao e um registro completo de todas as revisoes.

O relatorio apresenta pedidos observados, uso do historico, aceitacao,
correcoes, dependencia do algoritmo, divergencias, pedidos por perfil fisico
e frequencia de planos corrigidos. A contagem por perfil usa as atribuicoes
existentes; a auditoria decide quais referencias sao reutilizaveis.

Os registros antigos nao permitem reconstruir taxas de correcao anteriores.
O comparativo antes/depois exige coleta prospectiva e periodos comparaveis.
Perfis do historico sao calculados com o cadastro e regras atuais; nao ha
snapshot das medidas antigas. Distribuicoes legadas sem atribuicoes nao sao
classificadas como confirmacoes humanas. Este trabalho nao reinicia o servidor
nem modifica pedidos expedidos.
