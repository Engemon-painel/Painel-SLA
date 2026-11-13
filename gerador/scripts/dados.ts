// Office Script — gera o JSON do Painel-SLA a partir da aba "Historico"
//
// REGRAS:
// 1. Exclui completamente os incidentes da lista EXCLUIR_INCIDENTES.
//    Eles NÃO entram em Total, Dentro, Fora, diário, mensal, semanal ou FCA.
//
// 2. Os 3 incidentes da lista INCIDENTES_AGOSTO serão considerados
//    como pertencentes a AGOSTO, mesmo que a CLEAR_DATE seja de JULHO.
//
// 3. A CLEAR_DATE original NÃO é alterada.
//    Apenas o mês/data de referência dos indicadores é ajustado.
//
// 4. Fato, Causa e Ação vêm do FCA.xlsx através do fcaJson.
//
// 5. O incidente INC000005740882 está marcado como "Fora" na base,
//    mas deve ser tratado como "Dentro" do prazo (Eser 4h).
//
// 6. ALTERADO: o MTTR não vem mais de uma coluna "MTTR" pronta na aba
//    "Historico" — agora é calculado dentro do script, incidente por
//    incidente: SUBMIT_DATE até CLEAR_DATE e, se o CLEAR_DATE estiver
//    vazio/inválido, até o RESOLVED_DATE. O resultado (em segundos) é
//    exportado em "mttrValores" por mês, igual antes. Incidentes sem
//    SUBMIT_DATE nem uma data de fim válida ficam como null e são
//    ignorados no cálculo da média feito pelo painel.
//
// 7. NOVO: exporta também "horasFora" por mês — a hora (0-23) do
//    SUBMIT_DATE (não do CLEAR_DATE) de cada incidente "Fora" do prazo,
//    calculada separadamente para SLA 4h e 6h. Serve pra descobrir em
//    qual horário do dia mais se perde o prazo.
//
// 8. NOVO: os incidentes INC000005694775 (SPC2) e INC000005754202 (SPC3)
//    são forçados como "Dentro" tanto no ESER 4h quanto no ESER 6h, e
//    têm a AREA forçada, independente do que estiver na planilha.
//
// 9. NOVO: lê as colunas "GENERIC_CATEGORIZATION_TIER_2",
//    "GENERIC_CATEGORIZATION_TIER_3" e "RESOLUTION_CATEGORY_TIER_1" da aba
//    "Historico" e exporta "genericTier2Top10", "genericTier3Top10" e
//    "resolucaoTop10" DENTRO de cada mês (igual "topSites"): o Top 10 de
//    ocorrências de cada uma, só com os incidentes daquele mês. As três
//    colunas são lidas separadamente, sem fallback entre elas.
//
// 10. NOVO: lê a coluna "Unidade de Negócio" e a coluna "Coordenador" das
//     abas "Organograma"/"PJ" e exporta "unidadeNegocio" e "coordenador"
//     em cada colaborador, para o painel agrupar o Headcount por unidade
//     de negócio, exibindo o coordenador de cada unidade direto da coluna
//     (não mais o supervisor mais frequente).
//
// 11. NOVO: lê também as colunas "Matricula", "Nome", "Placa" e
//     "Desc.Funcao" das abas "Organograma"/"PJ" e exporta "matricula",
//     "nome", "placa" e "descFuncao" em cada colaborador, para a tabela
//     de busca de Colaboradores do Painel-SLA. Todas opcionais — se
//     alguma coluna não existir na planilha, o campo sai como string
//     vazia para todo mundo, sem quebrar o script.
//
// 12. NOVO: recebe também o parâmetro "rateioJson" — dados da aba "Rateio"
//     de OUTRA pasta de trabalho (lidos separadamente pelo Power Automate,
//     igual já acontece com o FCA.xlsx), no formato
//     [{ "placa": "ABC1D23", "valorTotal": 1234.56 }, ...]. Usado para
//     exportar "valorCombustivel" em cada veículo de "frota". Se o
//     parâmetro vier vazio ou a placa não aparecer no Rateio, o campo sai
//     como null, sem quebrar o script.
//
// 13. NOVO: recebe também o parâmetro "bancoHorasJson" — dados da aba
//     "qResumoFuncionario" de OUTRA pasta de trabalho
//     ("Banco_de_Horas_V2_Ajustes"), lidos separadamente pelo Power
//     Automate através de um script auxiliar (mesmo esquema do Rateio de
//     combustível) e repassados prontos pro painel em "bancoHoras". Se o
//     parâmetro vier vazio ou inválido, exporta lista vazia sem quebrar
//     o script.
//
// 14. NOVO: exporta também "diarioAreas" DENTRO de cada mês (tanto em
//     "meses" quanto em "meses6h") — quebra diária por área
//     (dia -> área -> {dentro, fora}), calculada a partir da MESMA
//     dataRef (CLEAR_DATE, caindo pro RESOLVED_DATE quando vazio) já
//     usada no resto do script. Vem direto da base "Historico", sem
//     depender do FCA. Usado pelo painel no gráfico "Diário por Área —
//     Fora do Prazo".
//
// 15. NOVO: inclui manualmente 7 incidentes de Setembro/2026
//     (INCIDENTES_MANUAL_SETEMBRO) que ainda não constavam na aba
//     "Historico" no momento da geração deste script. Eles entram em
//     Total, Dentro/Fora, MTTR, diário e diarioAreas normalmente, mas
//     sem SITE_CI/SUBMIT_DATE/CLEAR_DATE reais (não entram em Top
//     Sites, horasSubmit/horasClear nem nas colunas de Tier/Resolução).
//     Remover esse bloco quando esses incidentes passarem a existir de
//     verdade na base, pra não contar em duplicidade.
//
// 16. ALTERADO (NOVO): os incidentes de INCIDENTES_MANUAL_SETEMBRO agora
//     buscam Site/Submit/Clear/Resolved/Reported diretamente na aba
//     "Historico" pelo número do incidente (INCIDENT_NUMBER), através de
//     "dadosBrutosPorIncidente" — que indexa TODAS as linhas da aba, sem
//     aplicar os filtros de Área/Tipo/data usados no loop principal. Assim,
//     assim que o incidente aparecer na base (mesmo que a linha ainda não
//     passe em todos os filtros), os campos deixam de sair em branco no
//     Detalhamento — Fora do Prazo.
//     Além disso, se o incidente já tiver sido processado normalmente pelo
//     loop principal (ou seja, já existe "de verdade" na base e passa em
//     todos os filtros), o bloco manual é ignorado por completo pra esse
//     incidente — evita duplicidade, e a linha real da base assume os
//     dados e a classificação Dentro/Fora corretos.
//
// 17. NOVO: aba "Fila" (histórico de status por incidente: INCIDENT_NUMBER,
//     HIST_CREATE_DATE, HIST_INCIDENT_STATUS, HIST_ASSIGNED_SUPPORT_GROUP
//     entre outras). Para cada incidente "Fora" do prazo (tanto os que vêm
//     do loop principal quanto os de INCIDENTES_MANUAL_SETEMBRO), o script
//     cruza com a aba "Fila" pelo INCIDENT_NUMBER e calcula o TEMPO TOTAL
//     que o incidente ficou na fila: da primeira até a última linha de
//     HIST_CREATE_DATE daquele incidente (sem olhar o status). Exporta em
//     cada item de "pendentesFca": "tempoFilaSegundos" (número, ou null se
//     o incidente não aparecer na aba Fila) e "tempoFilaFormatado" (texto
//     tipo "2d 05h 30m", ou "" se não houver dado).
//
// 18. NOVO: exporta "porCidade" DENTRO de cada mês (tanto em "meses"
//     quanto em "meses6h") — agregação por Cidade (coluna "Cidade" da
//     aba "Historico"), no mesmo padrão de "topSites": para cada cidade
//     do mês, exporta qtd de incidentes, quantidade de sites únicos,
//     MTTR médio (em horas) e % dentro do prazo (eser4h/eser6h, conforme
//     o campo passado pra buildMeses). Usado pela tabela "Detalhe por
//     Cidade" na página "Por Site" do painel.
//
// 19. CORRIGIDO (10/09/2026): o bloco de INCIDENTES_MANUAL_SETEMBRO tinha
//     causa/acao/pendente FIXOS ("", "", true) e nunca consultava o
//     fcaPorIncidente — então mesmo com o Forms preenchido e o FCA.xlsx
//     atualizado, esses incidentes ficavam pra sempre como "A preencher"
//     no Detalhamento — Fora do Prazo. Agora esse bloco busca causa/acao
//     no fcaPorIncidente pelo número do incidente, igual o loop principal,
//     e "pendente" passa a ser calculado como causa === "".
//
// 20. NOVO: aba "Afaz" (Atividades de Falha — manutenção corretiva aberta
//     por site). Considera apenas linhas com Tipo = "Indisponibilidade" e
//     Status diferente de "Finalizado" (ou seja, ainda aberta). Indexa por
//     Site. Para cada incidente da aba "Historico", exporta o booleano
//     "teveAfazAberto": true se existir uma atividade de Afaz aberta
//     naquele mesmo Site, cuja "Abertura" seja ANTERIOR à ocorrência do
//     incidente (SUBMIT_DATE do incidente; se vazio, cai pro CLEAR/
//     RESOLVED usado como dataRef). Serve pra indicar que o site já tinha
//     um problema de manutenção em aberto quando o incidente aconteceu.
//
// 21. NOVO: aba "Cronograma" (agenda de preventivas por site). Considera
//     apenas linhas com Tipo = "Indisponibilidade". Indexa por Site.
//     Para cada incidente, exporta o booleano "tevePreventiva": true se
//     existir uma atividade de Cronograma naquele mesmo Site, cujo
//     "Início" seja ANTERIOR à ocorrência do incidente (mesma referência
//     de data usada no item 20). Serve pra indicar que houve uma
//     preventiva no site antes do incidente acontecer.
//     Ambos "teveAfazAberto" e "tevePreventiva" são exportados em cada
//     item de "pendentesFca" (Detalhamento — Fora do Prazo), inclusive
//     para os incidentes de INCIDENTES_MANUAL_SETEMBRO (usando o Site
//     encontrado via "dadosBrutosPorIncidente", quando disponível).
//
// 22. NOVO: exporta "incidentesPorSite" — lista simples (um item por
//     incidente de Indisponibilidade, Dentro e Fora) com Site, Incidente,
//     Área, Status (Eser 4h), mês (mesOrd) e dois campos de TEXTO (não
//     booleano): "afaz" com o valor da coluna "Área/Sistema" do Afaz
//     aberto mais recente encontrado (ou "" se não houver) e "preventiva"
//     com o valor da coluna "Tipo Preventiva" do Cronograma mais recente
//     encontrado (ou "" se não houver), ambos anteriores à ocorrência do
//     incidente. Usado pela tabela "Incidentes por Site" da página "Por
//     Site" do painel.
//
// 23. NOVO: exporta "acaoInterna" em cada item de "pendentesFca" — é a
//     coluna "AÇÃO" (H) do FCA.xlsx (valores como "Reclassificação"/
//     "Ação interna"), que já chega pelo campo "acao" existente no
//     fcaJson. Não depende de nenhuma mudança no fluxo do Power
//     Automate.
//
// 24. NOVO: exporta "acaoTomada" em cada item de "pendentesFca" — a
//     frase longa da coluna "Ação tomada" do FCA.xlsx (ex.: "Falha
//     tratada pelo OMR sem passagem na caixa de Infra."). Esse texto
//     NUNCA foi enviado pelo fluxo do Power Automate até agora — o
//     fcaJson precisa passar a incluir um campo "acaoTomada" por
//     incidente (lendo essa coluna do FCA.xlsx), senão sai como "" pra
//     todo mundo. Usado pela tabela "Causa x Ação Interna" da página
//     FCA, que agrupa os incidentes por Causa + Ação Interna e sugere
//     um plano de ação com base nos textos de "Ação tomada" mais
//     repetidos em cada grupo.
//
// 25. NOVO: lê também as colunas "TIPO DE INFRA" e "UNIDADE_RESPONSAVEL2"
//     da aba "Historico" e exporta "tipoInfra" e "unidadeResponsavel2"
//     em cada item de "incidentesPorSite" (mesma lista usada na tabela
//     "Incidentes por Site" da página "Por Site"). Usado pela nova aba
//     "Infra" do painel, que separa os incidentes Dentro/Fora do prazo
//     mostrando o Tipo de Infra e a Unidade Responsável 2 de cada um.
//     Para os incidentes de INCIDENTES_MANUAL_SETEMBRO, esses dois
//     campos saem vazios (sem dado bruto confiável pra eles).
//
// 26. NOVO: exporta "todosSites" DENTRO de cada mês (tanto em "meses"
//     quanto em "meses6h") — TODOS os sites do mês com a quantidade de
//     incidentes, sem o corte em 15 que existe em "topSites". Usado pelo
//     gráfico "Sites que Mais Repetiram" (página "Por Site"), que soma os
//     últimos 3 meses e precisa enxergar também os sites que ficaram fora
//     do Top 15 de algum mês.
//
// 27. NOVO: exporta "dadosAte" (dd/mm/aaaa) — data do incidente mais
//     recente encontrado na aba "Historico" (maior dataRef entre as
//     linhas). Diferente de "atualizado_em", que é só o dia em que este
//     script rodou. O painel mostra "base até dd/mm/aaaa" no cabeçalho e
//     avisa quando a base está com mais de 2 dias de atraso.
//
// 28. CORRIGIDO (21/09/2026): getISOWeek() usava a fórmula com o sinal
//     invertido ("- weekday + 3" em vez de "- 3 + weekday") e numerava as
//     semanas UMA a menos que o padrão ISO em 2025-2028 (ex.: 14/09/2026
//     saía como "Semana 37", mas é a semana ISO 38). Por isso a "Visão
//     semanal" do painel terminava na semana 37 quando o correto era 38
//     (a semana 39 começa em 21/09/2026). Agora a numeração é ISO 8601
//     (semana começa na segunda-feira), igual ao WEEKNUM(data;21) do Excel.
//
// 29. ALTERADO (21/09/2026): a "Visão semanal" (semanal/semanal6h) passa a
//     usar o número da coluna "Semana" da aba "Historico", em vez de
//     calcular a semana no script. Se a célula da coluna estiver vazia, ou
//     for um dos incidentes manuais de Setembro, usa a semana ISO calculada
//     da dataRef. Os 3 incidentes de INCIDENTES_AGOSTO continuam
//     contando pela dataRef ajustada (Agosto), como na regra 2. As semanas
//     são ordenadas pela data (não pelo número). O JSON traz "semanaFonte"
//     dizendo quantas linhas usaram a coluna.
//     >>> DESFEITO no item 43 (04/10/2026): a semana volta a ser calculada
//     pela dataRef (CLEAR_DATE, senão RESOLVED_DATE).
//
// 30. NOVO (21/09/2026): exporta "sitesBase" — contagem de incidentes por
//     SITE_CI direto das linhas da aba "Historico" (por mês), apenas com
//     Tipo = "Indisponibilidade" e sem filtrar por Área; também ignora
//     EXCLUIR_INCIDENTES.
//     O mês vem da primeira data preenchida (CLEAR, RESOLVED, SUBMIT ou
//     REPORTED). "sitesBaseSemData" traz as linhas com site mas sem nenhuma
//     data. Usado pelo ranking "Top Sites — Claro e Engemon" da página
//     "Por Site".
//     >>> ALTERADO no item 43 (04/10/2026): o mês vem só da CLEAR_DATE ou
//     da RESOLVED_DATE; sem nenhuma das duas, a linha vai pra
//     "sitesBaseSemData".
//
// 31. NOVO (21/09/2026): lê as colunas "Endereço" e "Num.Endereço" das abas
//     "Organograma"/"PJ" e exporta "endereco" (os dois juntos, ex.: "Rua X, 123")
//     em cada colaborador, para a coluna "Endereço" da tabela Colaboradores
//     do painel. Opcionais — se as colunas não existirem, sai "" pra todos.
//
// 32. NOVO (21/09/2026): removidos do Total e do Fora do Prazo os incidentes
//     INC000005849368, INC000005851458, INC000005851753, INC000005849357,
//     INC000005849825 e INC000005849794 — foram acrescentados a
//     EXCLUIR_INCIDENTES, e o bloco de INCIDENTES_MANUAL_SETEMBRO agora
//     também respeita essa lista (antes não checava). Eram 6 dos 7
//     incidentes manuais de setembro; o único que ficou é INC000005851798
//     ("Dentro").
//     >>> DESFEITO no item 42 (02/10/2026): esses 6 voltaram a contar.
//
// 33. NOVO (22/09/2026): escreve as justificativas do FCA.xlsx (Causa e Ação)
//     de volta na própria aba "Historico", em duas colunas novas — "Causa
//     (FCA)" e "Ação (FCA)" — cruzando pelo INCIDENT_NUMBER. Cria as colunas
//     na primeira vez (no fim da planilha) e, nas rodadas seguintes,
//     reaproveita as mesmas colunas em vez de duplicar. Incidentes sem
//     entrada no FCA ficam com as células vazias.
//
// 34. NOVO (22/09/2026): a mesma escrita de volta na aba "Historico" (item 33)
//     ganhou uma terceira coluna — "Ação tomada (FCA)" — com o texto longo da
//     coluna "Ação tomada" do FCA.xlsx (campo acaoTomada do fcaJson).
//
// 35. NOVO (22/09/2026): exporta "debugFcaMatch" (temporário) — mostra quantos
//     incidentes do fcaJson foram recebidos, quantos incidentes únicos existem
//     na aba "Historico", quantos deram match entre os dois, e uma amostra dos
//     que não bateram de cada lado. Serve pra diagnosticar por que as colunas
//     "Causa (FCA)"/"Ação (FCA)"/"Ação tomada (FCA)" (itens 33/34) saem vazias.
//
// 36. ALTERADO (23/09/2026): a leitura do fcaJson passa a aceitar os nomes dos
//     campos sem depender de maiúsculas, acentos ou espaços ("causa", "Causa",
//     "CAUSA", "Ação", "Ação tomada", "acaoTomada"...). Se o mesmo incidente
//     vier repetido no FCA, uma linha com causa vazia não apaga mais uma linha
//     preenchida. As colunas de Causa/Ação que já existem na base "Historico"
//     continuam sendo ignoradas — o script só lê do fcaJson e só grava nas
//     colunas "(FCA)". Exporta "debugFcaCampos" (temporário) com os nomes dos
//     campos que chegaram no fcaJson e quantos itens vieram com causa preenchida.
//
// 37. ALTERADO (23/09/2026): se a aba "Historico" tiver mais de uma coluna com o
//     mesmo nome ("Causa (FCA)", "Ação (FCA)" ou "Ação tomada (FCA)"), o script
//     grava em TODAS elas. Exporta "debugEscritaFca" (temporário) com as letras
//     das colunas gravadas, quantas linhas receberam causa e uma célula de
//     exemplo (endereço + texto) pra conferir direto no Excel.
//
// 38. ALTERADO (23/09/2026): o INCIDENT_NUMBER da aba "Historico" é ligado ao
//     "Incidente" do FCA também pelos dígitos do número (sem "INC", espaços ou
//     zeros à esquerda), além da comparação normal — vale pras colunas "(FCA)"
//     da Historico e pro Detalhamento — Fora do Prazo. "debugEscritaFca" ganha
//     "celulasPreenchidasAposGravar": relê a coluna logo depois de gravar.
//
// 39. CORRIGIDO (23/09/2026): quando o script era executado direto no Excel
//     (Automatizar → Executar), o fcaJson chegava vazio e o script GRAVAVA
//     células em branco por cima das colunas "(FCA)" — apagando o que o fluxo
//     tinha preenchido. Agora, se nenhum FCA chegar, as colunas não são
//     tocadas. A aba "FCA_Cruzado" (item 38) foi removida: tudo fica na aba
//     "Historico", e a aba antiga é apagada se ainda existir.
//
// 40. NOVO (28/09/2026): removidos os incidentes INC000005948061 e
//     INC000005948123 — acrescentados a EXCLUIR_INCIDENTES. Não entram em
//     Total, Dentro/Fora, MTTR, diário, semanal, mensal, FCA nem sitesBase.
//
// 41. NOVO (29/09/2026): recebe também o parâmetro "turnoverJson" — dados da
//     aba "Base" de OUTRA pasta de trabalho ("Analise de Turnover"), lidos
//     pelo Power Automate através do script auxiliar
//     "Turnover_LerBase" (mesmo esquema do Rateio e do Banco de Horas) e
//     repassados prontos pro painel em "turnover". Cada item traz
//     dataAdmis, dataDemissao (aaaa-mm-dd), supervisor, tipoDesligamento,
//     causaPareto e oQueEntra. Se o parâmetro vier vazio ou inválido,
//     exporta lista vazia sem quebrar o script.
//
// 42. ALTERADO (02/10/2026): os 6 incidentes do item 32 (INC000005849368,
//     INC000005851458, INC000005851753, INC000005849357, INC000005849825 e
//     INC000005849794) foram RETIRADOS de EXCLUIR_INCIDENTES e voltam a
//     contar em Total, Dentro/Fora, MTTR, diário, semanal, mensal, FCA e
//     sitesBase. Como já estão na aba "Historico", a linha real da base é
//     usada; o bloco INCIDENTES_MANUAL_SETEMBRO só entra como reserva se a
//     linha não passar nos filtros (Área SPC1-4 / Tipo / data) — e nunca
//     duplica (item 16).
//
// 43. ALTERADO (04/10/2026): TODO incidente passa a contar no mês (e na
//     semana) da CLEAR_DATE; se ela estiver vazia, da RESOLVED_DATE.
//     Ex.: SUBMIT em setembro e CLEAR em outubro -> conta em OUTUBRO.
//     Mudanças:
//     a) "Visão semanal" (semanal/semanal6h): a semana é a ISO da dataRef,
//        não mais a coluna "Semana" da aba Historico (desfaz o item 29).
//     b) INCIDENTES_MANUAL_SETEMBRO: dataRef e mesOrd vêm da CLEAR/RESOLVED
//        real do incidente (via dadosBrutosPorIncidente); só usam 01/09/2026
//        se o incidente não tiver nenhuma das duas datas na base.
//     c) "sitesBase": o mês vem só da CLEAR ou RESOLVED (não mais SUBMIT/
//        REPORTED); sem as duas, a linha vai pra "sitesBaseSemData".
//     A regra 2 (INCIDENTES_AGOSTO) continua valendo.
//
// 44. NOVO (05/10/2026): lê a coluna "MTTR" da aba "Historico" e exporta
//     "mttrHoras" em cada item de "incidentesPorSite". Se a célula estiver em
//     formato de hora (03:00), converte de dias para horas; se vier vazia, usa
//     o tempo calculado (SUBMIT até CLEAR/RESOLVED). Usado pelo quadro
//     "Sites com MTTR acima de 24h" da página "Por Site".
//
// 45. NOVO (06/10/2026): os 7 incidentes da lista INCIDENTES_OUTUBRO contam
//     em OUTUBRO/2026, mesmo os que têm CLEAR_DATE em 30/09/2026. Se a data
//     de referência (CLEAR, senão RESOLVED) cair antes de 01/10/2026, passa a
//     ser 01/10/2026 (mantendo a hora); se já for de outubro, fica como está.
//     Vale para mensal, diário, semanal, diarioAreas, FCA/Detalhamento,
//     incidentesPorSite, sitesBase e fcaTipoDivergente. A CLEAR_DATE original
//     NÃO é alterada (regra 3). Mesmo esquema da regra 2 (INCIDENTES_AGOSTO).
//
// 49. NOVO (12/11/2026): os 12 incidentes da lista INCIDENTES_NOVEMBRO contam
//     em NOVEMBRO/2026, mesmo com CLEAR_DATE em 31/10/2026 (mesmo esquema da
//     regra 45: a data de referência vira 01/11/2026 mantendo a hora; a
//     CLEAR_DATE original NÃO é alterada). Também contam como "Indisponibilidade"
//     e a Área vem de AREA_FORCADA_POR_INCIDENTE (SPC1/SPC2/SPC3/SPC4).
//
// 46. NOVO (06/10/2026): os incidentes de INCIDENTES_OUTUBRO são tratados
//     como Tipo = "Indisponibilidade" mesmo que a coluna "Tipo" da aba
//     "Historico" traga "Outros". A Área deles vem de AREA_FORCADA_POR_INCIDENTE
//     (SPC2/SPC3/SPC4, conforme a lista enviada). Também não aparecem mais em
//     "fcaTipoDivergente" e entram no "sitesBase".
//
// 47. NOVO (06/10/2026): INCIDENTES_MANUAL_OUTUBRO — os 3 incidentes do site
//     SMFCR13 (INC000006009954, INC000006009949, INC000006009864), SPC2,
//     Dentro no ESER 4h e 6h, com MTTR 02:52, 02:56 e 03:03. Eles entram de
//     duas formas, sem nunca duplicar:
//     a) se a linha existe na aba "Historico" e passa nos filtros, a linha real
//        é usada — e o status é forçado "Dentro" (4h e 6h), caso as colunas
//        Eser venham vazias ou diferentes;
//     b) se a linha não existe (ou não tem CLEAR/RESOLVED), entram pelo bloco
//        manual, igual INCIDENTES_MANUAL_SETEMBRO, com a data de limpeza
//        30/09/2026 ajustada para 01/10/2026 (regra 45).
//
// 48. NOVO (14/10/2026): aba "TOA" (atividades lançadas pelos técnicos no TOA).
//     Ignora as linhas com Tipo de Atividade = "Refeição". Agrupa por Data +
//     Recurso (técnico) e exporta "toa": [{ data (aaaa-mm-dd), recurso,
//     supervisor, total, concluidas }], onde "total" = atividades lançadas
//     (qualquer status) e "concluidas" = Status "Concluída". O supervisor vem
//     da coluna "Supervisor" (a última com esse nome na aba). Usado pela seção
//     "Produtividade TOA" do painel (gestao_toa.js): produtividade por técnico
//     por dia, por supervisor, e aderência ao TOA por supervisor. Exporta
//     "debugToa" (temporário) para conferir se a aba e as colunas foram lidas.
//     Se a aba não existir, "toa" sai como lista vazia sem quebrar o script.
//     ALTERADO (14/10/2026): as linhas de "Refeição" não são mais descartadas —
//     contam em "refeicao" (presença na base), mas continuam fora de "total" e
//     "concluidas". Cada item ganha "status": contagem por Status do TOA
//     (ex.: { "Concluída": 3, "Pendente": 1 }), sem as refeições.

function main(workbook: ExcelScript.Workbook, fcaJson: string, rateioJson: string, bancoHorasJson: string, turnoverJson: string): string {
  const NOME_ABA = "Historico";
  const AREAS_VALIDAS = ["SPC1", "SPC2", "SPC3", "SPC4"];
  const META = 0.85;
  const MES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];

  function normaliza(valor: unknown): string {
    return String(valor ?? "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\u00A0/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  const EXCLUIR_INCIDENTES: string[] = [
    "INC000005717988",
    "INC000005705794",
    "INC000005682547",
    "INC000005711476",
    "INC000005690201",
    "INC000005718033",
    "INC000005705910",
    // ALTERADO (item 42): os 6 incidentes do item 32 foram retirados desta
    // lista e voltaram a contar (INC000005849368, INC000005851458,
    // INC000005851753, INC000005849357, INC000005849825, INC000005849794).
    // NOVO (item 40): removidos de todos os indicadores.
    "INC000005948061",
    "INC000005948123"
  ];

  const INCIDENTES_AGOSTO: string[] = [
    "INC000005681055",
    "INC000005681028",
    "INC000005681008"
  ];

  // NOVO (item 45): incidentes que contam em OUTUBRO/2026.
  const INCIDENTES_OUTUBRO: string[] = [
    "INC000006010037",
    "INC000006009954",
    "INC000006009949",
    "INC000006009864",
    "INC000006009834",
    "INC000006009822",
    "INC000006009438"
  ];

  // NOVO (item 49, 12/11/2026): incidentes que contam em NOVEMBRO/2026 (CLEAR em 31/10).
  const INCIDENTES_NOVEMBRO: string[] = [
    "INC000006211062",
    "INC000006217380",
    "INC000006216156",
    "INC000006218621",
    "INC000006218739",
    "INC000006218528",
    "INC000006218490",
    "INC000006222762",
    "INC000006223294",
    "INC000006223396",
    "INC000006222764",
    "INC000006222630"
  ];

  const MTTR_MANUAL_POR_INCIDENTE: { [incidente: string]: number } = {
    "INC000005681055": 5580,
    "INC000005681028": 3180,
    "INC000005681008": 4500
  };

  const AREA_FORCADA_POR_INCIDENTE: { [incidente: string]: string } = {
    "INC000005681055": "SPC1",
    "INC000005681028": "SPC3",
    "INC000005681008": "SPC3",
    "INC000005694775": "SPC2",
    "INC000005754202": "SPC3",
    // NOVO (item 46): incidentes de outubro
    "INC000006010037": "SPC2",
    "INC000006009954": "SPC2",
    "INC000006009949": "SPC2",
    "INC000006009864": "SPC2",
    "INC000006009834": "SPC3",
    "INC000006009822": "SPC3",
    "INC000006009438": "SPC4",
    // NOVO (item 49): incidentes de novembro
    "INC000006211062": "SPC4",
    "INC000006217380": "SPC4",
    "INC000006216156": "SPC1",
    "INC000006218621": "SPC1",
    "INC000006218739": "SPC1",
    "INC000006218528": "SPC1",
    "INC000006218490": "SPC1",
    "INC000006222762": "SPC1",
    "INC000006223294": "SPC3",
    "INC000006223396": "SPC1",
    "INC000006222764": "SPC2",
    "INC000006222630": "SPC2"
  };

  const INCIDENTES_FORCAR_DENTRO: string[] = [
    "INC000005740882",
    "INC000005694775",
    "INC000005754202",
    // NOVO (item 47)
    "INC000006009954",
    "INC000006009949",
    "INC000006009864"
  ];

  const INCIDENTES_FORCAR_DENTRO_6H: string[] = [
    "INC000005694775",
    "INC000005754202",
    // NOVO (item 47)
    "INC000006009954",
    "INC000006009949",
    "INC000006009864"
  ];

  const incidentesExcluidos = new Set(EXCLUIR_INCIDENTES.map(inc => normaliza(inc)));
  const incidentesAgosto = new Set(INCIDENTES_AGOSTO.map(inc => normaliza(inc)));
  const incidentesOutubro = new Set(INCIDENTES_OUTUBRO.map(inc => normaliza(inc))); // NOVO (item 45)
  const incidentesNovembro = new Set(INCIDENTES_NOVEMBRO.map(inc => normaliza(inc))); // NOVO (item 49)
  const incidentesForcarDentro = new Set(INCIDENTES_FORCAR_DENTRO.map(inc => normaliza(inc)));
  const incidentesForcarDentro6h = new Set(INCIDENTES_FORCAR_DENTRO_6H.map(inc => normaliza(inc)));

  // NOVO (item 45): se a data cair antes de 01/10/2026, vira 01/10/2026
  // (mantendo a hora); se já for de outubro em diante, fica como está.
  const INICIO_OUTUBRO = new Date(2026, 9, 1);
  function ajustarParaOutubro(d: Date): Date {
    if (d.getTime() >= INICIO_OUTUBRO.getTime()) {
      return d;
    }
    return new Date(2026, 9, 1, d.getHours(), d.getMinutes(), d.getSeconds());
  }

  // NOVO (item 49): se a data cair antes de 01/11/2026, vira 01/11/2026 (mantendo a hora).
  const INICIO_NOVEMBRO = new Date(2026, 10, 1);
  function ajustarParaNovembro(d: Date): Date {
    if (d.getTime() >= INICIO_NOVEMBRO.getTime()) {
      return d;
    }
    return new Date(2026, 10, 1, d.getHours(), d.getMinutes(), d.getSeconds());
  }

  // NOVO (item 46): Tipo considerado pelo script. Os incidentes de
  // INCIDENTES_OUTUBRO contam como "Indisponibilidade" mesmo com "Outros" na base.
  function tipoEfetivo(tipoBase: unknown, incidente: unknown): string {
    const incN = normaliza(String(incidente ?? "").trim());
    if (incidentesOutubro.has(incN) || incidentesNovembro.has(incN)) {   // item 46 + item 49
      return "Indisponibilidade";
    }
    return String(tipoBase || "").trim();
  }

  const areaForcadaPorIncidente: { [incidenteNormalizado: string]: string } = {};
  Object.keys(AREA_FORCADA_POR_INCIDENTE).forEach(inc => {
    areaForcadaPorIncidente[normaliza(inc)] = AREA_FORCADA_POR_INCIDENTE[inc];
  });

  const mttrManualPorIncidente: { [incidenteNormalizado: string]: number } = {};
  Object.keys(MTTR_MANUAL_POR_INCIDENTE).forEach(inc => {
    mttrManualPorIncidente[normaliza(inc)] = MTTR_MANUAL_POR_INCIDENTE[inc];
  });

  const todasAbas = workbook.getWorksheets();
  let sheet = workbook.getWorksheet(NOME_ABA);
  if (!sheet) {
    const alvo = normaliza(NOME_ABA);
    for (const ws of todasAbas) {
      if (normaliza(ws.getName()) === alvo) {
        sheet = ws;
        break;
      }
    }
  }

  if (!sheet) {
    const nomes = todasAbas.map(ws => ws.getName()).join(", ");
    throw new Error("Aba 'Historico' não encontrada. Abas disponíveis: " + nomes);
  }

  const range = sheet.getUsedRange();
  if (!range) {
    throw new Error("A aba 'Historico' está vazia.");
  }

  const rowCount = range.getRowCount();
  const colCount = range.getColumnCount();

  if (rowCount < 2) {
    throw new Error("A aba 'Historico' não possui dados.");
  }

  const headers = sheet.getRangeByIndexes(0, 0, 1, colCount).getValues()[0];

  const colIndex: { [key: string]: number } = {};
  headers.forEach((h, i) => {
    const chave = normaliza(h);
    if (chave !== "") {
      colIndex[chave] = i;
    }
  });

  function getIdx(nomeColuna: string): number | undefined {
    const chave = normaliza(nomeColuna);
    return colIndex[chave];
  }

  const idxArea = getIdx("AREA");
  const idxTipo = getIdx("Tipo");
  const idxClear = getIdx("CLEAR_DATE");
  const idxResolved = getIdx("RESOLVED_DATE");
  const idxEser4h = getIdx("Eser 4h");
  const idxEser6h = getIdx("ESER 6H");
  const idxSupervisor = getIdx("Supervisor");
  const idxIncidente = getIdx("INCIDENT_NUMBER");
  const idxSite = getIdx("SITE_CI");
  const idxReported = getIdx("REPORTED_DATE");
  const idxSubmit = getIdx("SUBMIT_DATE");

  const idxGenericTier1 = getIdx("GENERIC_CATEGORIZATION_TIER_1"); // NOVO (11/11/2026) – Acompanhamento de Alarmes
  const idxGenericTier2 = getIdx("GENERIC_CATEGORIZATION_TIER_2");
  const idxGenericTier3 = getIdx("GENERIC_CATEGORIZATION_TIER_3");
  const idxResolucao = getIdx("RESOLUTION_CATEGORY_TIER_1");
  const idxCidade = getIdx("Cidade"); // NOVO
  const idxTipoInfra = getIdx("TIPO DE INFRA"); // NOVO (item 25)
  const idxUnidadeResp2 = getIdx("UNIDADE_RESPONSAVEL2"); // NOVO (item 25)
  const idxMttrBase = getIdx("MTTR") ?? getIdx("MTTR(h)") ?? getIdx("MTTR (h)"); // NOVO (item 44)

  const colunasEssenciais: { [key: string]: number | undefined } = {
    AREA: idxArea,
    Tipo: idxTipo,
    CLEAR_DATE: idxClear,
    "Eser 4h": idxEser4h
  };

  const faltando = Object.keys(colunasEssenciais).filter(k => colunasEssenciais[k] === undefined);

  if (faltando.length > 0) {
    throw new Error(
      "Colunas não encontradas: " + faltando.join(", ") +
      ". Cabeçalhos disponíveis: " + headers.map(h => String(h)).join(" | ")
    );
  }

  // ALTERADO (item 36): lê os campos do fcaJson sem depender de maiúsculas,
  // acentos ou espaços ("causa", "Causa", "CAUSA", "Ação tomada", "acaoTomada"...).
  // Se o mesmo incidente vier repetido, uma linha vazia não apaga uma preenchida.
  // "Ação Interna" (item 23) continua sendo o mesmo valor do campo "acao".
  function campoFca(item: { [k: string]: unknown }, nomes: string[]): string {
    const alvos = nomes.map(n => normaliza(n));
    for (const k of Object.keys(item)) {
      if (alvos.indexOf(normaliza(k)) !== -1) {
        const v = item[k];
        return v === null || v === undefined ? "" : String(v).trim();
      }
    }
    return "";
  }

  const fcaPorIncidente: { [incidente: string]: { fato: string; causa: string; acao: string; acaoInterna: string; acaoTomada: string; } } = {};
  let fcaChavesRecebidas: string[] = [];
  let fcaItensComCausa = 0;
  try {
    const fcaArray = JSON.parse(fcaJson || "[]") as { [k: string]: unknown }[];
    if (fcaArray.length > 0) {
      fcaChavesRecebidas = Object.keys(fcaArray[0]);
    }
    fcaArray.forEach(item => {
      const chave = normaliza(campoFca(item, ["incidente", "INCIDENT_NUMBER", "numero do incidente"]));
      if (chave === "") {
        return;
      }
      const acao = campoFca(item, ["acao"]);
      const registro = {
        fato: campoFca(item, ["fato"]),
        causa: campoFca(item, ["causa"]),
        acao: acao,
        acaoInterna: acao,
        acaoTomada: campoFca(item, ["acaoTomada", "acao tomada"])
      };
      if (registro.causa !== "") {
        fcaItensComCausa++;
      }
      const existente = fcaPorIncidente[chave];
      if (existente && existente.causa !== "" && registro.causa === "") {
        return; // não deixa uma linha vazia apagar uma preenchida
      }
      fcaPorIncidente[chave] = registro;
    });
  } catch (e) {
    // Se FCA estiver vazio/inválido, segue normalmente.
  }

  const rateioPorPlaca: { [placa: string]: number } = {};
  try {
    const rateioArray = JSON.parse(rateioJson || "[]") as { placa: string; valorTotal: number | string; }[];
    rateioArray.forEach(item => {
      const chave = normaliza(item.placa);
      if (chave === "") {
        return;
      }
      const valorBruto = item.valorTotal;
      const valorNum = typeof valorBruto === "number" ? valorBruto : Number(String(valorBruto || "").replace(/\./g, "").replace(",", "."));
      if (!isNaN(valorNum)) {
        rateioPorPlaca[chave] = (rateioPorPlaca[chave] || 0) + valorNum;
      }
    });
  } catch (e) {
    // Se Rateio estiver vazio/inválido, segue normalmente.
  }

  // NOVO (item 17): aba "Fila" — histórico de status por incidente. Usada só
  // pra calcular o tempo TOTAL que o incidente ficou na fila: da primeira
  // até a última linha de HIST_CREATE_DATE daquele incidente, sem olhar o
  // status. Cruzado depois com os incidentes "Fora" do prazo em Historico.
  interface FilaAgregada {
    primeira: number;
    ultima: number;
    primeiraAtribuido: number | null;
    primeiraResolvido: number | null;
    primeiraEncerrado: number | null;
    // NOVO: data mais antiga em que a linha da Fila foi criada pelo
    // sistema automático (PEOPLE_FULL_NAME = "App TemIP") — usada como
    // marco inicial pros tempos até Atribuído/Resolvido/Encerrado, no
    // lugar da primeira linha genérica.
    primeiraAppTemip: number | null;
    // NOVO: nomes (PEOPLE_FULL_NAME) de todas as pessoas que mexeram no
    // incidente, sem duplicar — chave normalizada, valor com o nome
    // original (como está na planilha).
    pessoas: { [nomeNormalizado: string]: string };
  }
  const filaPorIncidente: { [incidenteNormalizado: string]: FilaAgregada } = {};

  // NOVO (temporário): informações de diagnóstico sobre a leitura da aba
  // "Fila", exportadas no JSON final em "debugFila" — só pra investigar
  // por que os campos ficam vazios no painel. Pode ser removido depois
  // que o problema for identificado.
  interface DebugFilaAmostra {
    incidenteOriginal: string;
    dataOriginalBruta: string;
    dataSerialCalculada: number | null;
    statusOriginal: string;
    pessoaOriginal: string;
  }
  const debugFila: {
    abaEncontrada: boolean;
    colunasEncontradas: { incidente: boolean; data: boolean; status: boolean; pessoa: boolean };
    totalLinhasNaAba: number;
    totalLinhasComDataValida: number;
    totalIncidentesDistintos: number;
    amostraLinhas: DebugFilaAmostra[];
  } = {
    abaEncontrada: false,
    colunasEncontradas: { incidente: false, data: false, status: false, pessoa: false },
    totalLinhasNaAba: 0,
    totalLinhasComDataValida: 0,
    totalIncidentesDistintos: 0,
    amostraLinhas: []
  };

  const STATUS_KEYS_ATRIBUIDO = new Set(["assigned", "atribuido"]);
  const STATUS_KEYS_RESOLVIDO = new Set(["resolved", "resolvido"]);
  const STATUS_KEYS_ENCERRADO = new Set(["closed", "encerrado"]);
  const PESSOA_SISTEMA = "app temip"; // NOVO

  function menorData(atual: number | null, novo: number): number {
    return atual === null || novo < atual ? novo : atual;
  }

  function processarAbaFila(): void {
    const NOME_ABA_FILA = "Fila";
    let sheetFila = workbook.getWorksheet(NOME_ABA_FILA);
    if (!sheetFila) {
      const alvoFila = normaliza(NOME_ABA_FILA);
      for (const ws of todasAbas) {
        if (normaliza(ws.getName()) === alvoFila) {
          sheetFila = ws;
          break;
        }
      }
    }
    if (!sheetFila) {
      return;
    }
    debugFila.abaEncontrada = true;
    const rangeFila = sheetFila.getUsedRange();
    if (!rangeFila) {
      return;
    }
    const rowCountFila = rangeFila.getRowCount();
    const colCountFila = rangeFila.getColumnCount();
    if (rowCountFila < 2) {
      return;
    }
    const headersFila = sheetFila.getRangeByIndexes(0, 0, 1, colCountFila).getValues()[0];
    const colIndexFila: { [key: string]: number } = {};
    headersFila.forEach((h, i) => {
      const chave = normaliza(h);
      if (chave !== "") {
        colIndexFila[chave] = i;
      }
    });
    function getIdxFila(nomeColuna: string): number | undefined {
      return colIndexFila[normaliza(nomeColuna)];
    }
    const idxIncidenteFila = getIdxFila("INCIDENT_NUMBER");
    const idxDataFila = getIdxFila("HIST_CREATE_DATE");
    const idxStatusFila = getIdxFila("HIST_INCIDENT_STATUS");
    const idxPessoaFila = getIdxFila("PEOPLE_FULL_NAME"); // NOVO
    debugFila.colunasEncontradas = {
      incidente: idxIncidenteFila !== undefined,
      data: idxDataFila !== undefined,
      status: idxStatusFila !== undefined,
      pessoa: idxPessoaFila !== undefined
    };
    debugFila.totalLinhasNaAba = rowCountFila - 1;
    if (idxIncidenteFila === undefined || idxDataFila === undefined) {
      return;
    }

    function getColunaFila(idx: number): (string | number)[] {
      return sheetFila!.getRangeByIndexes(1, idx, rowCountFila - 1, 1).getValues().map(r => r[0]) as (string | number)[];
    }

    // NOVO: aceita a data tanto como número (serial do Excel) quanto como
    // texto no formato "dd/mm/aaaa hh:mm" ou "dd/mm/aaaa hh:mm:ss" —
    // algumas fontes colam a data como texto em vez de um valor de data
    // de verdade, e nesse caso o script precisa converter manualmente
    // pro serial equivalente pra poder comparar/calcular diferenças.
    function paraSerialFila(valor: string | number): number | null {
      if (typeof valor === "number" && valor > 0) {
        return valor;
      }
      const texto = String(valor || "").trim();
      if (texto === "") {
        return null;
      }
      const m = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?$/);
      if (!m) {
        return null;
      }
      const dia = Number(m[1]);
      const mes = Number(m[2]);
      const ano = Number(m[3]);
      const hora = Number(m[4]);
      const minuto = Number(m[5]);
      const segundo = m[6] !== undefined ? Number(m[6]) : 0;
      const d = new Date(ano, mes - 1, dia, hora, minuto, segundo);
      if (isNaN(d.getTime())) {
        return null;
      }
      const epoch = new Date(1899, 11, 30).getTime();
      return (d.getTime() - epoch) / 86400000;
    }

    const colIncidenteFila = getColunaFila(idxIncidenteFila);
    const colDataFila = getColunaFila(idxDataFila);
    const colStatusFila = idxStatusFila !== undefined ? getColunaFila(idxStatusFila) : [];
    const colPessoaFila = idxPessoaFila !== undefined ? getColunaFila(idxPessoaFila) : []; // NOVO

    let totalComDataValida = 0;

    for (let r = 0; r < rowCountFila - 1; r++) {
      const incNorm = normaliza(String(colIncidenteFila[r] || "").trim());
      const dataSerial = paraSerialFila(colDataFila[r]);

      // NOVO: guarda uma amostra das primeiras 5 linhas (processadas ou
      // não) pra diagnóstico — mostra exatamente o que o script está
      // vendo em cada coluna.
      if (debugFila.amostraLinhas.length < 5) {
        debugFila.amostraLinhas.push({
          incidenteOriginal: String(colIncidenteFila[r] || ""),
          dataOriginalBruta: String(colDataFila[r] || "") + " (tipo: " + typeof colDataFila[r] + ")",
          dataSerialCalculada: dataSerial,
          statusOriginal: idxStatusFila !== undefined ? String(colStatusFila[r] || "") : "(coluna não encontrada)",
          pessoaOriginal: idxPessoaFila !== undefined ? String(colPessoaFila[r] || "") : "(coluna não encontrada)"
        });
      }

      if (incNorm === "" || dataSerial === null || dataSerial <= 0) {
        continue;
      }
      totalComDataValida++;
      if (!filaPorIncidente[incNorm]) {
        filaPorIncidente[incNorm] = {
          primeira: dataSerial,
          ultima: dataSerial,
          primeiraAtribuido: null,
          primeiraResolvido: null,
          primeiraEncerrado: null,
          primeiraAppTemip: null,
          pessoas: {}
        };
      } else {
        if (dataSerial < filaPorIncidente[incNorm].primeira) {
          filaPorIncidente[incNorm].primeira = dataSerial;
        }
        if (dataSerial > filaPorIncidente[incNorm].ultima) {
          filaPorIncidente[incNorm].ultima = dataSerial;
        }
      }

      const statusNorm = idxStatusFila !== undefined ? normaliza(colStatusFila[r]) : "";
      if (STATUS_KEYS_ATRIBUIDO.has(statusNorm)) {
        filaPorIncidente[incNorm].primeiraAtribuido = menorData(filaPorIncidente[incNorm].primeiraAtribuido, dataSerial);
      } else if (STATUS_KEYS_RESOLVIDO.has(statusNorm)) {
        filaPorIncidente[incNorm].primeiraResolvido = menorData(filaPorIncidente[incNorm].primeiraResolvido, dataSerial);
      } else if (STATUS_KEYS_ENCERRADO.has(statusNorm)) {
        filaPorIncidente[incNorm].primeiraEncerrado = menorData(filaPorIncidente[incNorm].primeiraEncerrado, dataSerial);
      }

      // NOVO: registra a pessoa que mexeu no incidente nessa linha, e
      // marca a data mais antiga em que o sistema "App TemIP" criou
      // uma linha pra esse incidente.
      const pessoaBruta = idxPessoaFila !== undefined ? String(colPessoaFila[r] || "").trim() : "";
      if (pessoaBruta !== "") {
        const pessoaNorm = normaliza(pessoaBruta);
        filaPorIncidente[incNorm].pessoas[pessoaNorm] = pessoaBruta;
        if (pessoaNorm === PESSOA_SISTEMA) {
          filaPorIncidente[incNorm].primeiraAppTemip = menorData(filaPorIncidente[incNorm].primeiraAppTemip, dataSerial);
        }
      }
    }
    debugFila.totalLinhasComDataValida = totalComDataValida;
    debugFila.totalIncidentesDistintos = Object.keys(filaPorIncidente).length;
  }

  processarAbaFila();

  // ==================================================================
  // NOVO (item 20): aba "Afaz" — atividades de manutenção corretiva
  // abertas por Site. Só interessam linhas com Tipo = "Indisponibilidade"
  // e Status diferente de "Finalizado" (ou seja, ainda em aberto).
  // Indexado por Site normalizado, guardando a data de "Abertura" de
  // cada atividade aberta encontrada.
  // ==================================================================
  interface AfazAtividade {
    abertura: number | null;
    areaSistema: string; // NOVO (item 22): valor da coluna "Área/Sistema" do Afaz
  }
  const afazPorSite: { [siteNormalizado: string]: AfazAtividade[] } = {};

  function processarAbaAfaz(): void {
    const NOME_ABA_AFAZ = "Afaz";
    let sheetAfaz = workbook.getWorksheet(NOME_ABA_AFAZ);
    if (!sheetAfaz) {
      const alvoAfaz = normaliza(NOME_ABA_AFAZ);
      for (const ws of todasAbas) {
        if (normaliza(ws.getName()) === alvoAfaz) {
          sheetAfaz = ws;
          break;
        }
      }
    }
    if (!sheetAfaz) {
      return;
    }
    const rangeAfaz = sheetAfaz.getUsedRange();
    if (!rangeAfaz) {
      return;
    }
    const rowCountAfaz = rangeAfaz.getRowCount();
    const colCountAfaz = rangeAfaz.getColumnCount();
    if (rowCountAfaz < 2) {
      return;
    }
    const headersAfaz = sheetAfaz.getRangeByIndexes(0, 0, 1, colCountAfaz).getValues()[0];
    const colIndexAfaz: { [key: string]: number } = {};
    headersAfaz.forEach((h, i) => {
      const chave = normaliza(h);
      if (chave !== "") {
        colIndexAfaz[chave] = i;
      }
    });
    function getIdxAfaz(nomeColuna: string): number | undefined {
      return colIndexAfaz[normaliza(nomeColuna)];
    }
    const idxSiteAfaz = getIdxAfaz("Site");
    const idxStatusAfaz = getIdxAfaz("Status");
    const idxTipoAfaz = getIdxAfaz("Tipo");
    const idxAberturaAfaz = getIdxAfaz("Abertura");
    const idxAreaSistemaAfaz = getIdxAfaz("Área/Sistema"); // NOVO (item 22)

    if (idxSiteAfaz === undefined || idxTipoAfaz === undefined) {
      return;
    }

    function getColunaAfaz(idx: number | undefined): (string | number)[] {
      if (idx === undefined) {
        return [];
      }
      return sheetAfaz!.getRangeByIndexes(1, idx, rowCountAfaz - 1, 1).getValues().map(r => r[0]) as (string | number)[];
    }

    const colSiteAfaz = getColunaAfaz(idxSiteAfaz);
    const colStatusAfaz = getColunaAfaz(idxStatusAfaz);
    const colTipoAfaz = getColunaAfaz(idxTipoAfaz);
    const colAberturaAfaz = getColunaAfaz(idxAberturaAfaz);
    const colAreaSistemaAfaz = getColunaAfaz(idxAreaSistemaAfaz); // NOVO (item 22)

    for (let r = 0; r < rowCountAfaz - 1; r++) {
      const tipoVal = String(colTipoAfaz[r] || "").trim();
      if (tipoVal !== "Indisponibilidade") {
        continue;
      }
      const statusVal = String(colStatusAfaz[r] || "").trim();
      if (normaliza(statusVal) === normaliza("Finalizado")) {
        continue; // só interessa Afaz ainda aberto
      }
      const siteVal = String(colSiteAfaz[r] || "").trim();
      const siteNorm = normaliza(siteVal);
      if (siteNorm === "") {
        continue;
      }
      const aberturaBruta = colAberturaAfaz[r];
      const aberturaSerial = (typeof aberturaBruta === "number" && aberturaBruta > 0) ? aberturaBruta : null;
      const areaSistemaVal = String(colAreaSistemaAfaz[r] || "").trim(); // NOVO (item 22)
      if (!afazPorSite[siteNorm]) {
        afazPorSite[siteNorm] = [];
      }
      afazPorSite[siteNorm].push({ abertura: aberturaSerial, areaSistema: areaSistemaVal });
    }
  }

  processarAbaAfaz();

  // ==================================================================
  // NOVO (item 21): aba "Cronograma" — agenda de preventivas por Site.
  // Só interessam linhas com Tipo = "Indisponibilidade". Indexado por
  // Site normalizado, guardando a data de "Início" de cada atividade.
  // ==================================================================
  interface CronogramaAtividade {
    inicio: number | null;
    tipoPreventiva: string; // NOVO (item 22): valor da coluna "Tipo Preventiva" do Cronograma
  }
  const cronogramaPorSite: { [siteNormalizado: string]: CronogramaAtividade[] } = {};

  function processarAbaCronograma(): void {
    const NOME_ABA_CRONOGRAMA = "Cronograma";
    let sheetCronograma = workbook.getWorksheet(NOME_ABA_CRONOGRAMA);
    if (!sheetCronograma) {
      const alvoCron = normaliza(NOME_ABA_CRONOGRAMA);
      for (const ws of todasAbas) {
        if (normaliza(ws.getName()) === alvoCron) {
          sheetCronograma = ws;
          break;
        }
      }
    }
    if (!sheetCronograma) {
      return;
    }
    const rangeCron = sheetCronograma.getUsedRange();
    if (!rangeCron) {
      return;
    }
    const rowCountCron = rangeCron.getRowCount();
    const colCountCron = rangeCron.getColumnCount();
    if (rowCountCron < 2) {
      return;
    }
    const headersCron = sheetCronograma.getRangeByIndexes(0, 0, 1, colCountCron).getValues()[0];
    const colIndexCron: { [key: string]: number } = {};
    headersCron.forEach((h, i) => {
      const chave = normaliza(h);
      if (chave !== "") {
        colIndexCron[chave] = i;
      }
    });
    function getIdxCron(nomeColuna: string): number | undefined {
      return colIndexCron[normaliza(nomeColuna)];
    }
    const idxSiteCron = getIdxCron("Site");
    const idxTipoCron = getIdxCron("Tipo");
    const idxInicioCron = getIdxCron("Início");
    const idxTipoPreventivaCron = getIdxCron("Tipo Preventiva"); // NOVO (item 22)

    if (idxSiteCron === undefined || idxTipoCron === undefined) {
      return;
    }

    function getColunaCron(idx: number | undefined): (string | number)[] {
      if (idx === undefined) {
        return [];
      }
      return sheetCronograma!.getRangeByIndexes(1, idx, rowCountCron - 1, 1).getValues().map(r => r[0]) as (string | number)[];
    }

    const colSiteCron = getColunaCron(idxSiteCron);
    const colTipoCron = getColunaCron(idxTipoCron);
    const colInicioCron = getColunaCron(idxInicioCron);
    const colTipoPreventivaCron = getColunaCron(idxTipoPreventivaCron); // NOVO (item 22)

    for (let r = 0; r < rowCountCron - 1; r++) {
      const tipoVal = String(colTipoCron[r] || "").trim();
      if (tipoVal !== "Indisponibilidade") {
        continue;
      }
      const siteVal = String(colSiteCron[r] || "").trim();
      const siteNorm = normaliza(siteVal);
      if (siteNorm === "") {
        continue;
      }
      const inicioBruto = colInicioCron[r];
      const inicioSerial = (typeof inicioBruto === "number" && inicioBruto > 0) ? inicioBruto : null;
      const tipoPreventivaVal = String(colTipoPreventivaCron[r] || "").trim(); // NOVO (item 22)
      if (!cronogramaPorSite[siteNorm]) {
        cronogramaPorSite[siteNorm] = [];
      }
      cronogramaPorSite[siteNorm].push({ inicio: inicioSerial, tipoPreventiva: tipoPreventivaVal });
    }
  }

  processarAbaCronograma();

  function formatarDuracao(segundos: number | null): string {
    if (segundos === null || isNaN(segundos)) {
      return "";
    }
    const dias = Math.floor(segundos / 86400);
    const horas = Math.floor((segundos % 86400) / 3600);
    const minutos = Math.floor((segundos % 3600) / 60);
    if (dias > 0) {
      return dias + "d " + pad2(horas) + "h " + pad2(minutos) + "m";
    }
    return pad2(horas) + "h " + pad2(minutos) + "m";
  }

  function tempoFilaDoIncidente(incidenteNormalizado: string): number | null {
    const agregado = filaPorIncidente[incidenteNormalizado];
    if (!agregado) {
      return null;
    }
    const diffDias = agregado.ultima - agregado.primeira;
    if (diffDias < 0) {
      return null;
    }
    return Math.round(diffDias * 86400);
  }

  // ALTERADO: tempo entre o marco inicial do incidente e a primeira vez que
  // ele atingiu um status (Atribuído/Resolvido/Encerrado). O marco inicial
  // passa a ser a linha criada automaticamente pelo sistema "App TemIP"
  // (primeiraAppTemip); se o incidente não tiver nenhuma linha do App
  // TemIP na Fila, cai de volta pra primeira linha genérica (primeira).
  function tempoAteStatusDoIncidente(incidenteNormalizado: string, dataStatus: number | null): number | null {
    const agregado = filaPorIncidente[incidenteNormalizado];
    if (!agregado || dataStatus === null) {
      return null;
    }
    const origem = agregado.primeiraAppTemip !== null ? agregado.primeiraAppTemip : agregado.primeira;
    const diffDias = dataStatus - origem;
    if (diffDias < 0) {
      return null;
    }
    return Math.round(diffDias * 86400);
  }

  // NOVO: lista (sem duplicar) de todas as pessoas (PEOPLE_FULL_NAME) que
  // aparecem na Fila desse incidente.
  function quemMovimentouDoIncidente(incidenteNormalizado: string): string[] {
    const agregado = filaPorIncidente[incidenteNormalizado];
    if (!agregado) {
      return [];
    }
    return Object.keys(agregado.pessoas).map(k => agregado.pessoas[k]).sort();
  }

  // NOVO (item 20): true se existir uma atividade de "Afaz" aberta (Tipo =
  // Indisponibilidade, Status != Finalizado) no mesmo Site, com Abertura
  // ANTERIOR (ou igual) à data de referência do incidente.
  function teveAfazAbertoAntes(site: string, dataReferenciaSerial: number | null): boolean {
    if (dataReferenciaSerial === null) {
      return false;
    }
    const siteNorm = normaliza(site);
    if (siteNorm === "") {
      return false;
    }
    const lista = afazPorSite[siteNorm];
    if (!lista) {
      return false;
    }
    return lista.some(item => item.abertura !== null && item.abertura <= dataReferenciaSerial);
  }

  // NOVO (item 21): true se existir uma atividade de "Cronograma" (Tipo =
  // Indisponibilidade) no mesmo Site, com Início ANTERIOR (ou igual) à
  // data de referência do incidente.
  function tevePreventivaAntes(site: string, dataReferenciaSerial: number | null): boolean {
    if (dataReferenciaSerial === null) {
      return false;
    }
    const siteNorm = normaliza(site);
    if (siteNorm === "") {
      return false;
    }
    const lista = cronogramaPorSite[siteNorm];
    if (!lista) {
      return false;
    }
    return lista.some(item => item.inicio !== null && item.inicio <= dataReferenciaSerial);
  }

  // NOVO (item 22): retorna o valor da coluna "Área/Sistema" do Afaz aberto
  // mais RECENTE (maior Abertura) que seja anterior/igual à data de
  // referência, no mesmo Site — ou "" se não houver nenhum.
  function valorAfazAntes(site: string, dataReferenciaSerial: number | null): string {
    if (dataReferenciaSerial === null) {
      return "";
    }
    const siteNorm = normaliza(site);
    if (siteNorm === "") {
      return "";
    }
    const lista = afazPorSite[siteNorm];
    if (!lista) {
      return "";
    }
    let melhor: AfazAtividade | null = null;
    lista.forEach(item => {
      if (item.abertura !== null && item.abertura <= dataReferenciaSerial) {
        if (melhor === null || (item.abertura as number) > (melhor.abertura as number)) {
          melhor = item;
        }
      }
    });
    return melhor !== null ? (melhor as AfazAtividade).areaSistema : "";
  }

  // NOVO (item 22): retorna o valor da coluna "Tipo Preventiva" do
  // Cronograma mais RECENTE (maior Início) que seja anterior/igual à data
  // de referência, no mesmo Site — ou "" se não houver nenhum.
  function valorPreventivaAntes(site: string, dataReferenciaSerial: number | null): string {
    if (dataReferenciaSerial === null) {
      return "";
    }
    const siteNorm = normaliza(site);
    if (siteNorm === "") {
      return "";
    }
    const lista = cronogramaPorSite[siteNorm];
    if (!lista) {
      return "";
    }
    let melhor: CronogramaAtividade | null = null;
    lista.forEach(item => {
      if (item.inicio !== null && item.inicio <= dataReferenciaSerial) {
        if (melhor === null || (item.inicio as number) > (melhor.inicio as number)) {
          melhor = item;
        }
      }
    });
    return melhor !== null ? (melhor as CronogramaAtividade).tipoPreventiva : "";
  }

  function getColuna(idx: number | undefined): (string | number)[] {
    if (idx === undefined) {
      return [];
    }
    return sheet!.getRangeByIndexes(1, idx, rowCount - 1, 1).getValues().map(r => r[0]) as (string | number)[];
  }

  const colArea = getColuna(idxArea);
  const colTipo = getColuna(idxTipo);
  const colClear = getColuna(idxClear);
  const colResolved = getColuna(idxResolved);
  const colEser4h = getColuna(idxEser4h);
  const colEser6h = getColuna(idxEser6h);
  const colSupervisor = getColuna(idxSupervisor);
  const colIncidente = getColuna(idxIncidente);
  const colSite = getColuna(idxSite);
  const colReported = getColuna(idxReported);
  const colSubmit = getColuna(idxSubmit);

  const colGenericTier1 = getColuna(idxGenericTier1); // NOVO (11/11/2026)
  const colGenericTier2 = getColuna(idxGenericTier2);
  const colGenericTier3 = getColuna(idxGenericTier3);
  const colResolucao = getColuna(idxResolucao);
  const colCidade = getColuna(idxCidade); // NOVO
  const colTipoInfra = getColuna(idxTipoInfra); // NOVO (item 25)
  const colUnidadeResp2 = getColuna(idxUnidadeResp2); // NOVO (item 25)

  // NOVO (item 44): MTTR da coluna "MTTR" da base, em horas. Se a célula estiver
  // em formato de hora (ex.: 03:00 ou [h]:mm), o Excel guarda em dias -> converte x24.
  const colMttrBase = getColuna(idxMttrBase);
  const colMttrBaseTexto: string[] = idxMttrBase === undefined
    ? []
    : sheet!.getRangeByIndexes(1, idxMttrBase, rowCount - 1, 1).getTexts().map(r => String(r[0]));
  function mttrHorasDaBase(r: number): number | null {
    const v = colMttrBase[r];
    if (typeof v !== "number" || isNaN(v) || v < 0) {
      return null;
    }
    const txt = colMttrBaseTexto[r] || "";
    return txt.indexOf(":") !== -1 ? v * 24 : v;
  }

  // NOVO (item 33): escreve as justificativas do FCA.xlsx (Causa e Ação) DENTRO da
  // própria aba "Historico" do Excel, em duas colunas novas — "Causa (FCA)" e
  // "Ação (FCA)". Cruza pelo INCIDENT_NUMBER com o mesmo fcaPorIncidente já usado no
  // resto do script. Se as colunas já existirem (rodadas anteriores), reaproveita —
  // não duplica — e sobrescreve os valores, inclusive limpando quando o incidente não
  // tem mais entrada no FCA. Não altera nenhuma outra coluna nem a ordem das linhas.
  function localizarOuCriarColunaHistorico(nomeColuna: string): number {
    const chave = normaliza(nomeColuna);
    if (colIndex[chave] !== undefined) {
      return colIndex[chave];
    }
    const novoIndice = headers.length;
    sheet!.getRangeByIndexes(0, novoIndice, 1, 1).setValues([[nomeColuna]]);
    colIndex[chave] = novoIndice;
    headers.push(nomeColuna);
    return novoIndice;
  }

  // NOVO (item 35): diagnóstico do cruzamento — quantos incidentes do FCA.xlsx
  // (fcaJson) foram encontrados na aba "Historico" pelo INCIDENT_NUMBER.
  const debugFcaMatch: {
    totalFcaRecebidos: number;
    totalHistoricoIncidentesUnicos: number;
    totalComMatch: number;
    amostraFcaSemMatch: string[];
    amostraHistoricoSemFca: string[];
  } = {
    totalFcaRecebidos: 0,
    totalHistoricoIncidentesUnicos: 0,
    totalComMatch: 0,
    amostraFcaSemMatch: [],
    amostraHistoricoSemFca: []
  };
  function calcularDebugFcaMatch(): void {
    try {
      const fcaArrayDebug = JSON.parse(fcaJson || "[]") as unknown[];
      debugFcaMatch.totalFcaRecebidos = fcaArrayDebug.length;
    } catch (e) {
      debugFcaMatch.totalFcaRecebidos = 0;
    }
    const incidentesHistoricoUnicos = new Set<string>();
    for (let r = 0; r < rowCount - 1; r++) {
      const incBruto = String(colIncidente[r] || "").trim();
      if (incBruto !== "") {
        incidentesHistoricoUnicos.add(normaliza(incBruto));
      }
    }
    debugFcaMatch.totalHistoricoIncidentesUnicos = incidentesHistoricoUnicos.size;
    Object.keys(fcaPorIncidente).forEach(chave => {
      if (incidentesHistoricoUnicos.has(chave)) {
        debugFcaMatch.totalComMatch++;
      } else if (debugFcaMatch.amostraFcaSemMatch.length < 5) {
        debugFcaMatch.amostraFcaSemMatch.push(chave);
      }
    });
    incidentesHistoricoUnicos.forEach(chave => {
      if (fcaPorIncidente[chave] === undefined && debugFcaMatch.amostraHistoricoSemFca.length < 5) {
        debugFcaMatch.amostraHistoricoSemFca.push(chave);
      }
    });
  }
  calcularDebugFcaMatch();

  // ALTERADO (item 37): se existir mais de uma coluna com o mesmo nome
  // (ex.: duas "Causa (FCA)" criadas em rodadas diferentes), grava em TODAS.
  function letraColuna(idx: number): string {
    let s = "";
    let n = idx + 1;
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function indicesDaColunaHistorico(nomeColuna: string): number[] {
    const alvo = normaliza(nomeColuna);
    const encontrados: number[] = [];
    headers.forEach((h, i) => {
      if (normaliza(h) === alvo) {
        encontrados.push(i);
      }
    });
    if (encontrados.length === 0) {
      encontrados.push(localizarOuCriarColunaHistorico(nomeColuna));
    }
    return encontrados;
  }

  const debugEscritaFca: { colunasCausa: string[]; colunasAcao: string[]; colunasAcaoTomada: string[]; linhasComCausa: number; celulasPreenchidasAposGravar: number; exemploCelula: string; exemploTexto: string } = {
    colunasCausa: [],
    colunasAcao: [],
    colunasAcaoTomada: [],
    linhasComCausa: 0,
    celulasPreenchidasAposGravar: 0,
    exemploCelula: "",
    exemploTexto: ""
  };

  // NOVO (item 38): chave do incidente só com os dígitos, sem zeros à esquerda.
  function chaveDigitosIncidente(valor: unknown): string {
    return String(valor ?? "").replace(/\D/g, "").replace(/^0+/, "");
  }
  const fcaPorDigitos: { [digitos: string]: { fato: string; causa: string; acao: string; acaoInterna: string; acaoTomada: string; } } = {};
  Object.keys(fcaPorIncidente).forEach(chave => {
    const d = chaveDigitosIncidente(chave);
    if (d !== "") {
      fcaPorDigitos[d] = fcaPorIncidente[chave];
    }
  });
  function buscarFca(incidenteBruto: unknown) {
    const porTexto = fcaPorIncidente[normaliza(String(incidenteBruto ?? "").trim())];
    if (porTexto) {
      return porTexto;
    }
    const d = chaveDigitosIncidente(incidenteBruto);
    return d !== "" ? fcaPorDigitos[d] : undefined;
  }

  function escreverJustificativasNoHistorico(): void {
    const idxsCausa = indicesDaColunaHistorico("Causa (FCA)");
    const idxsAcao = indicesDaColunaHistorico("Ação (FCA)");
    const idxsAcaoTomada = indicesDaColunaHistorico("Ação tomada (FCA)"); // NOVO (item 34)

    const valoresCausa: string[][] = [];
    const valoresAcao: string[][] = [];
    const valoresAcaoTomada: string[][] = []; // NOVO (item 34)

    for (let r = 0; r < rowCount - 1; r++) {
      const incBruto = String(colIncidente[r] || "").trim();
      const fcaResp = incBruto !== "" ? buscarFca(incBruto) : undefined;
      const causa = fcaResp ? fcaResp.causa : "";
      const acao = fcaResp ? fcaResp.acao : "";
      const acaoTomada = fcaResp ? fcaResp.acaoTomada : "";
      valoresCausa.push([causa]);
      valoresAcao.push([acao]);
      valoresAcaoTomada.push([acaoTomada]); // NOVO (item 34)
      if (causa !== "") {
        debugEscritaFca.linhasComCausa++;
        if (debugEscritaFca.exemploCelula === "") {
          debugEscritaFca.exemploCelula = letraColuna(idxsCausa[0]) + (r + 2);
          debugEscritaFca.exemploTexto = causa;
        }
      }
    }

    idxsCausa.forEach(idx => {
      sheet!.getRangeByIndexes(1, idx, rowCount - 1, 1).setValues(valoresCausa);
      debugEscritaFca.colunasCausa.push(letraColuna(idx));
    });
    idxsAcao.forEach(idx => {
      sheet!.getRangeByIndexes(1, idx, rowCount - 1, 1).setValues(valoresAcao);
      debugEscritaFca.colunasAcao.push(letraColuna(idx));
    });
    idxsAcaoTomada.forEach(idx => {
      sheet!.getRangeByIndexes(1, idx, rowCount - 1, 1).setValues(valoresAcaoTomada);
      debugEscritaFca.colunasAcaoTomada.push(letraColuna(idx));
    });

    // NOVO (item 38): relê a coluna "Causa (FCA)" logo depois de gravar.
    const releitura = sheet!.getRangeByIndexes(1, idxsCausa[0], rowCount - 1, 1).getValues();
    releitura.forEach(linha => {
      if (String(linha[0] ?? "").trim() !== "") {
        debugEscritaFca.celulasPreenchidasAposGravar++;
      }
    });
  }

  // NOVO (item 39): só grava as colunas "(FCA)" quando o FCA realmente chegou.
  if (Object.keys(fcaPorIncidente).length > 0) {
    escreverJustificativasNoHistorico();
  }

  // NOVO (item 39): remove a aba antiga "FCA_Cruzado", se ainda existir.
  const abaCruzadoAntiga = workbook.getWorksheet("FCA_Cruzado");
  if (abaCruzadoAntiga) {
    abaCruzadoAntiga.delete();
  }

  function serialToDate(serial: number): Date {
    const utcDays = Math.floor(serial - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);
    const fractionalDay = serial - Math.floor(serial) + 0.0000001;
    let totalSeconds = Math.floor(86400 * fractionalDay);
    const seconds = totalSeconds % 60;
    totalSeconds -= seconds;
    const hours = Math.floor(totalSeconds / (60 * 60));
    const minutes = Math.floor(totalSeconds / 60) % 60;
    return new Date(dateInfo.getFullYear(), dateInfo.getMonth(), dateInfo.getDate(), hours, minutes, seconds);
  }

  function calcularMttrSegundos(submitSerial: number | null, clearSerial: number | null, resolvedSerial: number | null): number | null {
    if (submitSerial === null) {
      return null;
    }
    const fimSerial = clearSerial !== null ? clearSerial : resolvedSerial;
    if (fimSerial === null) {
      return null;
    }
    const diffDias = fimSerial - submitSerial;
    if (diffDias < 0) {
      return null;
    }
    return Math.round(diffDias * 86400);
  }

  // NOVO (11/11/2026): serial do Excel → "aaaa-mm-dd hh:mm" (Acompanhamento de Alarmes)
  function isoDoSerial(sd: number | null): string {
    if (sd === null) return "";
    const dd = serialToDate(sd);
    return dd.getFullYear() + "-" + pad2(dd.getMonth() + 1) + "-" + pad2(dd.getDate()) + " " + pad2(dd.getHours()) + ":" + pad2(dd.getMinutes());
  }
  function pad2(n: number): string {
    return n < 10 ? "0" + n : String(n);
  }

  function formatDateTime(serial: number | null | undefined): string {
    if (typeof serial !== "number" || serial <= 0) {
      return "";
    }
    const d = serialToDate(serial);
    return pad2(d.getDate()) + "/" + pad2(d.getMonth() + 1) + "/" + d.getFullYear() + " " + pad2(d.getHours()) + ":" + pad2(d.getMinutes()) + ":" + pad2(d.getSeconds());
  }

  // ISO 8601 (semana começa na segunda), igual ao WEEKNUM(data;21) do Excel.
  function getISOWeek(d: Date): number {
    const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    const dayNum = (date.getUTCDay() + 6) % 7;
    date.setUTCDate(date.getUTCDate() - dayNum + 3);
    const firstThursday = new Date(Date.UTC(date.getUTCFullYear(), 0, 4));
    const diff = (date.getTime() - firstThursday.getTime()) / 86400000;
    return 1 + Math.round((diff - 3 + ((firstThursday.getUTCDay() + 6) % 7)) / 7);
  }

  interface Linha {
    area: string;
    dataRef: Date;
    eser4h: string;
    eser6h: string;
    mttr: number | null;
    submitHora: number | null;
    // NOVO (item 19): hora (0-23) do REPORTED_DATE (horário do alarme).
    reportedHora: number | null;
    site: string;
    genericTier2: string;
    genericTier3: string;
    resolucao: string;
    cidade: string; // NOVO (item 18)
    teveAfazAberto: boolean; // NOVO (item 20)
    tevePreventiva: boolean; // NOVO (item 21)
  }

  interface Pendencia {
    supervisor: string;
    area: string;
    incidente: string;
    site: string;
    clear: string;
    resolved: string;
    reported: string;
    submit: string;
    mesOrd: string;
    causa: string;
    acao: string;
    acaoInterna: string; // NOVO (item 23)
    acaoTomada: string; // NOVO (item 24)
    pendente: boolean;
    tempoFilaSegundos: number | null;
    tempoFilaFormatado: string;
    tempoAteAtribuidoSegundos: number | null;
    tempoAteAtribuidoFormatado: string;
    tempoAteResolvidoSegundos: number | null;
    tempoAteResolvidoFormatado: string;
    tempoAteEncerradoSegundos: number | null;
    tempoAteEncerradoFormatado: string;
    quemMovimentou: string[];
    teveAfazAberto: boolean; // NOVO (item 20)
    tevePreventiva: boolean; // NOVO (item 21)
  }

  // NOVO (item 22): lista simples pra tabela "Incidentes por Site".
  interface IncidentePorSite {
    site: string;
    incidente: string;
    area: string;
    eser4h: string;
    mesOrd: string;
    afaz: string;
    preventiva: string;
    tipoInfra: string; // NOVO (item 25)
    unidadeResponsavel2: string; // NOVO (item 25)
    mttrHoras: number | null; // NOVO (item 44): MTTR em horas (coluna MTTR da base)
    data: string; // NOVO (11/11/2026): dia e hora da queda (SUBMIT_DATE) "aaaa-mm-dd hh:mm" – Acompanhamento de Alarmes
  }
  const incidentesPorSite: IncidentePorSite[] = [];

  const linhas: Linha[] = [];
  const pendentesFca: Pendencia[] = [];

  const contagemSupervisorPorArea: { [area: string]: { [supervisor: string]: number } } = {};

  // NOVO (item 16): índice bruto de TODAS as linhas da aba "Historico" por
  // número de incidente, sem nenhum filtro de Área/Tipo/data.
  interface DadosBrutosIncidente {
    site: string;
    submit: number | null;
    clear: number | null;
    resolved: number | null;
    reported: number | null;
    tipo: string;
    tipoBase: string; // NOVO (13/11/2026): tipo como está na base (antes de forçar Indisponibilidade)
    area: string;
    supervisor: string;
    incidenteOriginal: string;
    cidade: string; // NOVO (item 18)
    cat1: string; // NOVO (11/11/2026): GENERIC_CATEGORIZATION_TIER_1
    cat2: string; // NOVO (11/11/2026): GENERIC_CATEGORIZATION_TIER_2
    cat3: string; // NOVO (12/11/2026): GENERIC_CATEGORIZATION_TIER_3
  }
  const dadosBrutosPorIncidente: { [incidenteNormalizado: string]: DadosBrutosIncidente } = {};
  for (let r = 0; r < rowCount - 1; r++) {
    const incNumBruto = String(colIncidente[r] || "").trim();
    const incNormBruto = normaliza(incNumBruto);
    if (incNormBruto === "") {
      continue;
    }
    dadosBrutosPorIncidente[incNormBruto] = {
      site: String(colSite[r] || "").trim(),
      submit: (typeof colSubmit[r] === "number" && (colSubmit[r] as number) > 0) ? (colSubmit[r] as number) : null,
      clear: (typeof colClear[r] === "number" && (colClear[r] as number) > 0) ? (colClear[r] as number) : null,
      resolved: (typeof colResolved[r] === "number" && (colResolved[r] as number) > 0) ? (colResolved[r] as number) : null,
      reported: (typeof colReported[r] === "number" && (colReported[r] as number) > 0) ? (colReported[r] as number) : null,
      tipo: tipoEfetivo(colTipo[r], incNumBruto), // ALTERADO (item 46)
      tipoBase: String(colTipo[r] || "").trim(), // NOVO (13/11/2026)
      area: String(colArea[r] || "").trim(),
      supervisor: String(colSupervisor[r] || "").trim(),
      incidenteOriginal: incNumBruto,
      cidade: String(colCidade[r] || "").trim(), // NOVO (item 18)
      cat1: String(colGenericTier1[r] || "").trim(), // NOVO (11/11/2026)
      cat2: String(colGenericTier2[r] || "").trim(), // NOVO (11/11/2026)
      cat3: String(colGenericTier3[r] || "").trim() // NOVO (12/11/2026)
    };
  }

  // NOVO (11/11/2026) – ACOMPANHAMENTO DE ALARMES: todas as quedas da aba "Historico"
  // (Indisponibilidade E Outros), sem filtro de Área, a partir de 01/07/2026.
  // s = site · d = SUBMIT_DATE "aaaa-mm-dd hh:mm" · t = "I" (Indisponibilidade) ou "O" (outros) · i = incidente
  // c1 / c2 = GENERIC_CATEGORIZATION_TIER_1 / _TIER_2 (ex.: ENERGIA CONCESSIONARIA / FALHA ENERGIA)
  const quedasSite: { s: string; d: string; t: string; i: string; c1: string; c2: string; c3: string }[] = [];
  const INICIO_QUEDAS = new Date(2026, 6, 1).getTime();
  Object.keys(dadosBrutosPorIncidente).forEach(k => {
    const b = dadosBrutosPorIncidente[k];
    const sd = b.submit !== null ? b.submit : b.reported;
    if (sd === null || b.site === "") return;
    if (serialToDate(sd).getTime() < INICIO_QUEDAS) return;
    quedasSite.push({ s: b.site, d: isoDoSerial(sd), t: b.tipo === "Indisponibilidade" ? "I" : "O", i: b.incidenteOriginal, c1: b.cat1, c2: b.cat2, c3: b.cat3 });
  });

  // NOVO (item 17/18): incidentes justificados no FCA cujo Tipo ATUAL na
  // aba "Historico" não é mais "Indisponibilidade". mesOrd pela
  // CLEAR_DATE (ou RESOLVED_DATE).
  interface FcaTipoDivergente {
    incidente: string;
    area: string;
    supervisor: string;
    tipoAtual: string;
    causa: string;
    acao: string;
    mesOrd: string;
  }
  const fcaTipoDivergente: FcaTipoDivergente[] = [];
  Object.keys(fcaPorIncidente).forEach(chaveNormalizada => {
    const brutos = dadosBrutosPorIncidente[chaveNormalizada];
    if (!brutos || brutos.tipo === "") {
      return;
    }
    // NOVO (13/11/2026): os incidentes de INCIDENTES_NOVEMBRO contam no ESER de
    // novembro como Indisponibilidade E também aparecem nos reclassificados,
    // com o tipo que está na base.
    const tipoLista = incidentesNovembro.has(chaveNormalizada) ? brutos.tipoBase : brutos.tipo;
    if (tipoLista !== "" && tipoLista !== "Indisponibilidade") {
      const fcaResp = fcaPorIncidente[chaveNormalizada];
      const serialRef = brutos.clear !== null ? brutos.clear : brutos.resolved;
      let mesOrd = "";
      if (serialRef !== null) {
        let dRef = serialToDate(serialRef);
        if (incidentesOutubro.has(chaveNormalizada)) {
          dRef = ajustarParaOutubro(dRef); // NOVO (item 45)
        }
        if (incidentesNovembro.has(chaveNormalizada)) {
          dRef = ajustarParaNovembro(dRef); // NOVO (item 49)
        }
        mesOrd = dRef.getFullYear() + "-" + pad2(dRef.getMonth() + 1);
      }
      fcaTipoDivergente.push({
        incidente: brutos.incidenteOriginal,
        area: brutos.area,
        supervisor: brutos.supervisor,
        tipoAtual: tipoLista,
        causa: fcaResp.causa || "",
        acao: fcaResp.acao || "",
        mesOrd: mesOrd
      });
    }
  });

  // NOVO (item 16): incidentes que passaram de verdade pelo loop principal.
  const incidentesRealProcessados = new Set<string>();

  for (let r = 0; r < rowCount - 1; r++) {
    let area = String(colArea[r] || "").trim();
    const incidenteNum = String(colIncidente[r] || "").trim();
    const tipo = tipoEfetivo(colTipo[r], incidenteNum); // ALTERADO (item 46)
    const incidenteNormalizado = normaliza(incidenteNum);

    if (areaForcadaPorIncidente[incidenteNormalizado] !== undefined) {
      area = areaForcadaPorIncidente[incidenteNormalizado];
    }

    if (AREAS_VALIDAS.indexOf(area) === -1) {
      continue;
    }

    if (tipo !== "Indisponibilidade") {
      continue;
    }

    if (incidenteNormalizado !== "" && incidentesExcluidos.has(incidenteNormalizado)) {
      continue;
    }

    // Data de referência do incidente: CLEAR_DATE; se vazia, RESOLVED_DATE.
    // É ELA que define o mês, o dia e a semana em que o incidente conta.
    const clearVal = colClear[r];
    const resolvedVal = colResolved[r];
    let serial: number | null = null;
    if (typeof clearVal === "number" && clearVal > 0) {
      serial = clearVal;
    } else if (typeof resolvedVal === "number" && resolvedVal > 0) {
      serial = resolvedVal;
    }

    if (serial === null) {
      continue;
    }

    const dataOriginal = serialToDate(serial);
    let dataRef = dataOriginal;
    if (incidentesAgosto.has(incidenteNormalizado)) {
      dataRef = new Date(2026, 7, 1, dataOriginal.getHours(), dataOriginal.getMinutes(), dataOriginal.getSeconds());
    }
    // NOVO (item 45): incidentes que contam em outubro/2026.
    if (incidentesOutubro.has(incidenteNormalizado)) {
      dataRef = ajustarParaOutubro(dataRef);
    }
    // NOVO (item 49): incidentes que contam em novembro/2026.
    if (incidentesNovembro.has(incidenteNormalizado)) {
      dataRef = ajustarParaNovembro(dataRef);
    }

    let eser4h = String(colEser4h[r] || "").trim();
    let eser6h = String(colEser6h[r] || "").trim();

    if (incidentesForcarDentro.has(incidenteNormalizado)) {
      eser4h = "Dentro";
    }
    if (incidentesForcarDentro6h.has(incidenteNormalizado)) {
      eser6h = "Dentro";
    }

    const submitValBruto = colSubmit[r];
    const submitSerialLinha: number | null = (typeof submitValBruto === "number" && submitValBruto > 0) ? submitValBruto : null;
    const clearSerialLinha: number | null = (typeof colClear[r] === "number" && (colClear[r] as number) > 0) ? (colClear[r] as number) : null;
    const resolvedSerialLinha: number | null = (typeof colResolved[r] === "number" && (colResolved[r] as number) > 0) ? (colResolved[r] as number) : null;

    const mttrCalculado = calcularMttrSegundos(submitSerialLinha, clearSerialLinha, resolvedSerialLinha);
    const mttr = mttrManualPorIncidente[incidenteNormalizado] !== undefined ? mttrManualPorIncidente[incidenteNormalizado] : mttrCalculado;

    const submitHora: number | null = submitSerialLinha !== null ? serialToDate(submitSerialLinha).getHours() : null;

    // NOVO (item 19): hora do REPORTED_DATE (horário do alarme).
    const reportedValBruto = colReported[r];
    const reportedSerialLinha: number | null = (typeof reportedValBruto === "number" && reportedValBruto > 0) ? reportedValBruto : null;
    const reportedHora: number | null = reportedSerialLinha !== null ? serialToDate(reportedSerialLinha).getHours() : null;

    const genericTier2Val = String(colGenericTier2[r] || "").trim();
    const genericTier3Val = String(colGenericTier3[r] || "").trim();
    const resolucaoFinal = String(colResolucao[r] || "").trim();
    const cidadeVal = String(colCidade[r] || "").trim(); // NOVO (item 18)
    const siteDaLinha = String(colSite[r] || "").trim();

    // NOVO (item 20/21): data da OCORRÊNCIA (SUBMIT_DATE; se vazio, a dataRef).
    const dataOcorrenciaSerial: number | null = submitSerialLinha !== null ? submitSerialLinha : serial;
    const teveAfazAbertoVal = teveAfazAbertoAntes(siteDaLinha, dataOcorrenciaSerial);
    const tevePreventivaVal = tevePreventivaAntes(siteDaLinha, dataOcorrenciaSerial);

    // NOVO (item 22): valores (texto) de Afaz/Preventiva.
    const afazValorLinha = valorAfazAntes(siteDaLinha, dataOcorrenciaSerial);
    const preventivaValorLinha = valorPreventivaAntes(siteDaLinha, dataOcorrenciaSerial);
    const tipoInfraVal = String(colTipoInfra[r] || "").trim(); // NOVO (item 25)
    const unidadeResp2Val = String(colUnidadeResp2[r] || "").trim(); // NOVO (item 25)
    incidentesPorSite.push({
      site: siteDaLinha,
      incidente: incidenteNum,
      area: area,
      eser4h: eser4h,
      mesOrd: dataRef.getFullYear() + "-" + pad2(dataRef.getMonth() + 1),
      afaz: afazValorLinha,
      preventiva: preventivaValorLinha,
      tipoInfra: tipoInfraVal,
      unidadeResponsavel2: unidadeResp2Val,
      mttrHoras: mttrHorasDaBase(r) ?? (mttr !== null ? mttr / 3600 : null), // NOVO (item 44)
      data: isoDoSerial(submitSerialLinha !== null ? submitSerialLinha : (reportedSerialLinha !== null ? reportedSerialLinha : dataOcorrenciaSerial))
    });

    if (incidenteNormalizado !== "") {
      incidentesRealProcessados.add(incidenteNormalizado);
    }

    linhas.push({
      area: area,
      dataRef: dataRef,
      eser4h: eser4h,
      eser6h: eser6h,
      mttr: mttr,
      submitHora: submitHora,
      reportedHora: reportedHora,
      site: siteDaLinha,
      genericTier2: genericTier2Val,
      genericTier3: genericTier3Val,
      resolucao: resolucaoFinal,
      cidade: cidadeVal,
      teveAfazAberto: teveAfazAbertoVal,
      tevePreventiva: tevePreventivaVal
    });

    const supervisorDaLinha = String(colSupervisor[r] || "").trim();
    if (supervisorDaLinha !== "") {
      if (!contagemSupervisorPorArea[area]) {
        contagemSupervisorPorArea[area] = {};
      }
      contagemSupervisorPorArea[area][supervisorDaLinha] = (contagemSupervisorPorArea[area][supervisorDaLinha] || 0) + 1;
    }

    if (eser4h === "Fora") {
      const fcaResp = buscarFca(incidenteNum); // ALTERADO (item 38)
      const causa = fcaResp ? fcaResp.causa : "";
      const acao = fcaResp ? fcaResp.acao : "";
      const acaoInterna = fcaResp ? fcaResp.acaoInterna : ""; // NOVO (item 23)
      const acaoTomada = fcaResp ? fcaResp.acaoTomada : ""; // NOVO (item 24)
      const pendente = causa === "";
      const mesOrd = dataRef.getFullYear() + "-" + pad2(dataRef.getMonth() + 1);

      const tempoFilaSeg = tempoFilaDoIncidente(incidenteNormalizado);
      const agregadoFila = filaPorIncidente[incidenteNormalizado];
      const tempoAteAtribuidoSeg = agregadoFila ? tempoAteStatusDoIncidente(incidenteNormalizado, agregadoFila.primeiraAtribuido) : null;
      const tempoAteResolvidoSeg = agregadoFila ? tempoAteStatusDoIncidente(incidenteNormalizado, agregadoFila.primeiraResolvido) : null;
      const tempoAteEncerradoSeg = agregadoFila ? tempoAteStatusDoIncidente(incidenteNormalizado, agregadoFila.primeiraEncerrado) : null;
      const quemMovimentouLista = quemMovimentouDoIncidente(incidenteNormalizado);

      pendentesFca.push({
        supervisor: String(colSupervisor[r] || "").trim(),
        area: area,
        incidente: incidenteNum,
        site: siteDaLinha,
        clear: formatDateTime(colClear[r] as number),
        resolved: formatDateTime(colResolved[r] as number),
        reported: formatDateTime(colReported[r] as number),
        submit: formatDateTime(colSubmit[r] as number),
        mesOrd: mesOrd,
        causa: causa,
        acao: acao,
        acaoInterna: acaoInterna,
        acaoTomada: acaoTomada,
        pendente: pendente,
        tempoFilaSegundos: tempoFilaSeg,
        tempoFilaFormatado: formatarDuracao(tempoFilaSeg),
        tempoAteAtribuidoSegundos: tempoAteAtribuidoSeg,
        tempoAteAtribuidoFormatado: formatarDuracao(tempoAteAtribuidoSeg),
        tempoAteResolvidoSegundos: tempoAteResolvidoSeg,
        tempoAteResolvidoFormatado: formatarDuracao(tempoAteResolvidoSeg),
        tempoAteEncerradoSegundos: tempoAteEncerradoSeg,
        tempoAteEncerradoFormatado: formatarDuracao(tempoAteEncerradoSeg),
        quemMovimentou: quemMovimentouLista,
        teveAfazAberto: teveAfazAbertoVal,
        tevePreventiva: tevePreventivaVal
      });
    }
  }

  // ALTERADO (item 42): reserva — só entra se a linha real não passar nos filtros (item 16).
  interface IncidenteManual {
    incidente: string;
    area: string;
    eser4h: "Dentro" | "Fora";
    eser6h: "Dentro" | "Fora";
    mttrSegundos: number;
    site?: string;        // NOVO (item 47): usado se o incidente não estiver na base
    cidade?: string;      // NOVO (item 47)
    dataLimpeza?: Date;   // NOVO (item 47): data de referência se não houver CLEAR/RESOLVED na base
  }
  const INCIDENTES_MANUAL_SETEMBRO: IncidenteManual[] = [
    { incidente: "INC000005851798", area: "SPC1", eser4h: "Dentro", eser6h: "Dentro", mttrSegundos: 660 },
    { incidente: "INC000005851753", area: "SPC2", eser4h: "Fora", eser6h: "Fora", mttrSegundos: 33480 },
    { incidente: "INC000005851458", area: "SPC3", eser4h: "Fora", eser6h: "Fora", mttrSegundos: 32640 },
    { incidente: "INC000005849825", area: "SPC2", eser4h: "Fora", eser6h: "Fora", mttrSegundos: 70380 },
    { incidente: "INC000005849794", area: "SPC2", eser4h: "Fora", eser6h: "Fora", mttrSegundos: 70980 },
    { incidente: "INC000005849368", area: "SPC2", eser4h: "Fora", eser6h: "Fora", mttrSegundos: 55140 },
    { incidente: "INC000005849357", area: "SPC2", eser4h: "Fora", eser6h: "Fora", mttrSegundos: 70080 }
  ];

  // NOVO (item 47): reserva para os 3 incidentes do SMFCR13 (contam em outubro).
  const INCIDENTES_MANUAL_OUTUBRO: IncidenteManual[] = [
    {
      incidente: "INC000006009954", area: "SPC2", eser4h: "Dentro", eser6h: "Dentro", mttrSegundos: 10320,
      site: "SMFCR13", cidade: "FRANCO DA ROCHA", dataLimpeza: new Date(2026, 8, 30, 23, 49, 0)
    },
    {
      incidente: "INC000006009949", area: "SPC2", eser4h: "Dentro", eser6h: "Dentro", mttrSegundos: 10560,
      site: "SMFCR13", cidade: "FRANCO DA ROCHA", dataLimpeza: new Date(2026, 8, 30, 23, 51, 0)
    },
    {
      incidente: "INC000006009864", area: "SPC2", eser4h: "Dentro", eser6h: "Dentro", mttrSegundos: 10980,
      site: "SMFCR13", cidade: "FRANCO DA ROCHA", dataLimpeza: new Date(2026, 8, 30, 23, 49, 0)
    }
  ];

  const supervisorMaisComumPorArea: { [area: string]: string } = {};
  Object.keys(contagemSupervisorPorArea).forEach(area => {
    const contagem = contagemSupervisorPorArea[area];
    let melhorSupervisor = "";
    let melhorContagem = -1;
    Object.keys(contagem).forEach(supervisor => {
      if (contagem[supervisor] > melhorContagem) {
        melhorContagem = contagem[supervisor];
        melhorSupervisor = supervisor;
      }
    });
    supervisorMaisComumPorArea[area] = melhorSupervisor;
  });

  INCIDENTES_MANUAL_SETEMBRO.concat(INCIDENTES_MANUAL_OUTUBRO).forEach(item => { // ALTERADO (item 47)
    const incidenteNormalizadoManual = normaliza(item.incidente);

    if (incidentesRealProcessados.has(incidenteNormalizadoManual)) {
      return;
    }

    if (incidentesExcluidos.has(incidenteNormalizadoManual)) {
      return;
    }

    const brutos = dadosBrutosPorIncidente[incidenteNormalizadoManual];

    // ALTERADO (item 43): mês pela CLEAR_DATE real (ou RESOLVED_DATE);
    // só usa 01/09/2026 se o incidente não tiver nenhuma das duas na base.
    const serialRefManual: number | null = brutos ? (brutos.clear !== null ? brutos.clear : brutos.resolved) : null;
    let dataRefManual = serialRefManual !== null ? serialToDate(serialRefManual) : (item.dataLimpeza || new Date(2026, 8, 1));
    if (incidentesOutubro.has(incidenteNormalizadoManual)) {
      dataRefManual = ajustarParaOutubro(dataRefManual); // NOVO (item 47): regra 45
    }
    if (incidentesNovembro.has(incidenteNormalizadoManual)) {
      dataRefManual = ajustarParaNovembro(dataRefManual); // NOVO (item 49)
    }
    const mesOrdManual = dataRefManual.getFullYear() + "-" + pad2(dataRefManual.getMonth() + 1);

    const siteManual = brutos && brutos.site !== "" ? brutos.site : (item.site || ""); // ALTERADO (item 47)
    const dataOcorrenciaManualSerial: number | null = brutos
      ? (brutos.submit !== null ? brutos.submit : (brutos.clear !== null ? brutos.clear : brutos.resolved))
      : null;
    const teveAfazAbertoManual = teveAfazAbertoAntes(siteManual, dataOcorrenciaManualSerial);
    const tevePreventivaManual = tevePreventivaAntes(siteManual, dataOcorrenciaManualSerial);

    incidentesPorSite.push({
      site: siteManual,
      incidente: item.incidente,
      area: item.area,
      eser4h: item.eser4h,
      mesOrd: mesOrdManual, // ALTERADO (item 43)
      afaz: valorAfazAntes(siteManual, dataOcorrenciaManualSerial),
      preventiva: valorPreventivaAntes(siteManual, dataOcorrenciaManualSerial),
      tipoInfra: "",
      unidadeResponsavel2: "",
      mttrHoras: item.mttrSegundos / 3600, // NOVO (item 44)
      data: isoDoSerial(dataOcorrenciaManualSerial !== null ? dataOcorrenciaManualSerial : (brutos ? brutos.reported : null))
    });

    linhas.push({
      area: item.area,
      dataRef: dataRefManual,
      eser4h: item.eser4h,
      eser6h: item.eser6h,
      mttr: item.mttrSegundos,
      submitHora: null,
      reportedHora: null,
      site: siteManual,
      genericTier2: "",
      genericTier3: "",
      resolucao: "",
      cidade: brutos && brutos.cidade !== "" ? brutos.cidade : (item.cidade || ""), // ALTERADO (item 47)
      teveAfazAberto: teveAfazAbertoManual,
      tevePreventiva: tevePreventivaManual
    });

    if (item.eser4h === "Fora") {
      const tempoFilaSegManual = tempoFilaDoIncidente(incidenteNormalizadoManual);
      const agregadoFilaManual = filaPorIncidente[incidenteNormalizadoManual];
      const tempoAteAtribuidoSegM = agregadoFilaManual ? tempoAteStatusDoIncidente(incidenteNormalizadoManual, agregadoFilaManual.primeiraAtribuido) : null;
      const tempoAteResolvidoSegM = agregadoFilaManual ? tempoAteStatusDoIncidente(incidenteNormalizadoManual, agregadoFilaManual.primeiraResolvido) : null;
      const tempoAteEncerradoSegM = agregadoFilaManual ? tempoAteStatusDoIncidente(incidenteNormalizadoManual, agregadoFilaManual.primeiraEncerrado) : null;
      const quemMovimentouListaManual = quemMovimentouDoIncidente(incidenteNormalizadoManual);

      const fcaRespManual = buscarFca(item.incidente);
      const causaManual = fcaRespManual ? fcaRespManual.causa : "";
      const acaoManual = fcaRespManual ? fcaRespManual.acao : "";
      const acaoInternaManual = fcaRespManual ? fcaRespManual.acaoInterna : "";
      const acaoTomadaManual = fcaRespManual ? fcaRespManual.acaoTomada : "";

      pendentesFca.push({
        supervisor: supervisorMaisComumPorArea[item.area] || "",
        area: item.area,
        incidente: item.incidente,
        site: siteManual,
        clear: brutos ? formatDateTime(brutos.clear) : "",
        resolved: brutos ? formatDateTime(brutos.resolved) : "",
        reported: brutos ? formatDateTime(brutos.reported) : "",
        submit: brutos ? formatDateTime(brutos.submit) : "",
        mesOrd: mesOrdManual, // ALTERADO (item 43)
        causa: causaManual,
        acao: acaoManual,
        acaoInterna: acaoInternaManual,
        acaoTomada: acaoTomadaManual,
        pendente: causaManual === "",
        tempoFilaSegundos: tempoFilaSegManual,
        tempoFilaFormatado: formatarDuracao(tempoFilaSegManual),
        tempoAteAtribuidoSegundos: tempoAteAtribuidoSegM,
        tempoAteAtribuidoFormatado: formatarDuracao(tempoAteAtribuidoSegM),
        tempoAteResolvidoSegundos: tempoAteResolvidoSegM,
        tempoAteResolvidoFormatado: formatarDuracao(tempoAteResolvidoSegM),
        tempoAteEncerradoSegundos: tempoAteEncerradoSegM,
        tempoAteEncerradoFormatado: formatarDuracao(tempoAteEncerradoSegM),
        quemMovimentou: quemMovimentouListaManual,
        teveAfazAberto: teveAfazAbertoManual,
        tevePreventiva: tevePreventivaManual
      });
    }
  });

  function buildDiario(campo: "eser4h" | "eser6h") {
    const map: { [key: string]: { data: string; dentro: number; fora: number; ord: string; } } = {};
    linhas.forEach(l => {
      const val = l[campo];
      if (val !== "Dentro" && val !== "Fora") {
        return;
      }
      const dd = pad2(l.dataRef.getDate());
      const mm = pad2(l.dataRef.getMonth() + 1);
      const yyyy = l.dataRef.getFullYear();
      const dataStr = dd + "/" + mm;
      const ord = yyyy + "-" + mm + "-" + dd;
      if (!map[ord]) {
        map[ord] = { data: dataStr, dentro: 0, fora: 0, ord: ord };
      }
      if (val === "Dentro") {
        map[ord].dentro++;
      } else {
        map[ord].fora++;
      }
    });
    return Object.keys(map).sort().map(k => ({ data: map[k].data, dentro: map[k].dentro, fora: map[k].fora }));
  }

  const diario = buildDiario("eser4h");
  const diario6h = buildDiario("eser6h");

  function buildMttrPorMes(): { [ord: string]: (number | null)[] } {
    const map: { [ord: string]: (number | null)[] } = {};
    linhas.forEach(l => {
      const ord = l.dataRef.getFullYear() + "-" + pad2(l.dataRef.getMonth() + 1);
      if (!map[ord]) {
        map[ord] = [];
      }
      map[ord].push(l.mttr);
    });
    return map;
  }

  const mttrPorMes = buildMttrPorMes();

  function buildSitesPorMes(): { [ord: string]: { [site: string]: number } } {
    const map: { [ord: string]: { [site: string]: number } } = {};
    linhas.forEach(l => {
      const ord = l.dataRef.getFullYear() + "-" + pad2(l.dataRef.getMonth() + 1);
      if (!map[ord]) {
        map[ord] = {};
      }
      const site = l.site !== "" ? l.site : "Sem site";
      map[ord][site] = (map[ord][site] || 0) + 1;
    });
    return map;
  }

  const sitesPorMes = buildSitesPorMes();
  const TOP_SITES_LIMITE = 15;

  function topSitesDoMes(ord: string): { site: string; total: number }[] {
    const contagem = sitesPorMes[ord] || {};
    return Object.keys(contagem).map(site => ({ site: site, total: contagem[site] })).sort((a, b) => b.total - a.total).slice(0, TOP_SITES_LIMITE);
  }

  // NOVO (item 26): TODOS os sites do mês (sem o corte em 15).
  function todosSitesDoMes(ord: string): { site: string; total: number }[] {
    const contagem = sitesPorMes[ord] || {};
    return Object.keys(contagem)
      .map(site => ({ site: site, total: contagem[site] }))
      .sort((a, b) => b.total - a.total);
  }

  function totalSitesUnicosDoMes(ord: string): number {
    const contagem = sitesPorMes[ord] || {};
    return Object.keys(contagem).length;
  }

  function buildContagemPorMes(extrator: (l: Linha) => string): { [ord: string]: { [nome: string]: number } } {
    const map: { [ord: string]: { [nome: string]: number } } = {};
    linhas.forEach(l => {
      const valor = extrator(l);
      if (valor === "") {
        return;
      }
      const ord = l.dataRef.getFullYear() + "-" + pad2(l.dataRef.getMonth() + 1);
      if (!map[ord]) {
        map[ord] = {};
      }
      map[ord][valor] = (map[ord][valor] || 0) + 1;
    });
    return map;
  }

  const genericTier2PorMes = buildContagemPorMes(l => l.genericTier2);
  const genericTier3PorMes = buildContagemPorMes(l => l.genericTier3);
  const resolucaoPorMes = buildContagemPorMes(l => l.resolucao);
  const TOP_CATEGORIAS_LIMITE = 10;

  function topContagemDoMes(mapaPorMes: { [ord: string]: { [nome: string]: number } }, ord: string): { nome: string; total: number }[] {
    const contagem = mapaPorMes[ord] || {};
    return Object.keys(contagem).map(nome => ({ nome: nome, total: contagem[nome] })).sort((a, b) => b.total - a.total).slice(0, TOP_CATEGORIAS_LIMITE);
  }

  interface HoraStat {
    hora: string;
    total: number;
    fora: number;
  }

  function buildHorasStatsPorMes(campo: "eser4h" | "eser6h", fonte: "submit" | "clear" | "reported"): { [ord: string]: HoraStat[] } {
    const contagem: { [ord: string]: { total: number[]; fora: number[]; } } = {};
    linhas.forEach(l => {
      const val = l[campo];
      if (val !== "Dentro" && val !== "Fora") {
        return;
      }
      const hora: number | null = fonte === "submit" ? l.submitHora : (fonte === "reported" ? l.reportedHora : l.dataRef.getHours());
      if (hora === null) {
        return;
      }
      const ord = l.dataRef.getFullYear() + "-" + pad2(l.dataRef.getMonth() + 1);
      if (!contagem[ord]) {
        contagem[ord] = { total: new Array(24).fill(0), fora: new Array(24).fill(0) };
      }
      contagem[ord].total[hora]++;
      if (val === "Fora") {
        contagem[ord].fora[hora]++;
      }
    });
    const resultado: { [ord: string]: HoraStat[] } = {};
    Object.keys(contagem).forEach(ord => {
      resultado[ord] = [];
      for (let h = 0; h < 24; h++) {
        resultado[ord].push({ hora: pad2(h), total: contagem[ord].total[h], fora: contagem[ord].fora[h] });
      }
    });
    return resultado;
  }

  // ALTERADO (item 19): "horasSubmit" usa a hora do REPORTED_DATE (nome do campo mantido).
  const horasSubmitStatsPorMes4h = buildHorasStatsPorMes("eser4h", "reported");
  const horasSubmitStatsPorMes6h = buildHorasStatsPorMes("eser6h", "reported");
  const horasClearStatsPorMes4h = buildHorasStatsPorMes("eser4h", "clear");
  const horasClearStatsPorMes6h = buildHorasStatsPorMes("eser6h", "clear");

  function buildMeses(campo: "eser4h" | "eser6h") {
    const horasSubmitStatsPorMes = campo === "eser4h" ? horasSubmitStatsPorMes4h : horasSubmitStatsPorMes6h;
    const horasClearStatsPorMes = campo === "eser4h" ? horasClearStatsPorMes4h : horasClearStatsPorMes6h;
    const map: {
      [mesKey: string]: {
        mes: string;
        ord: string;
        areas: { [a: string]: { dentro: number; fora: number; } };
        diarioAreasMap: { [dia: string]: { [a: string]: { dentro: number; fora: number; } } };
        cidadesMap: { [cidade: string]: { dentro: number; fora: number; sites: Set<string>; mttrValores: number[]; } };
      }
    } = {};

    linhas.forEach(l => {
      const val = l[campo];
      if (val !== "Dentro" && val !== "Fora") {
        return;
      }
      const mesLabel = MES_PT[l.dataRef.getMonth()] + "/" + pad2(l.dataRef.getFullYear() % 100);
      const ord = l.dataRef.getFullYear() + "-" + pad2(l.dataRef.getMonth() + 1);
      if (!map[ord]) {
        map[ord] = { mes: mesLabel, ord: ord, areas: {}, diarioAreasMap: {}, cidadesMap: {} };
      }
      if (!map[ord].areas[l.area]) {
        map[ord].areas[l.area] = { dentro: 0, fora: 0 };
      }
      if (val === "Dentro") {
        map[ord].areas[l.area].dentro++;
      } else {
        map[ord].areas[l.area].fora++;
      }
      const diaRef = pad2(l.dataRef.getDate());
      if (!map[ord].diarioAreasMap[diaRef]) {
        map[ord].diarioAreasMap[diaRef] = {};
      }
      if (!map[ord].diarioAreasMap[diaRef][l.area]) {
        map[ord].diarioAreasMap[diaRef][l.area] = { dentro: 0, fora: 0 };
      }
      if (val === "Dentro") {
        map[ord].diarioAreasMap[diaRef][l.area].dentro++;
      } else {
        map[ord].diarioAreasMap[diaRef][l.area].fora++;
      }

      // NOVO (item 18): agregação por Cidade.
      const cidadeVal = l.cidade !== "" ? l.cidade : "Sem cidade";
      if (!map[ord].cidadesMap[cidadeVal]) {
        map[ord].cidadesMap[cidadeVal] = { dentro: 0, fora: 0, sites: new Set<string>(), mttrValores: [] };
      }
      if (val === "Dentro") {
        map[ord].cidadesMap[cidadeVal].dentro++;
      } else {
        map[ord].cidadesMap[cidadeVal].fora++;
      }
      if (l.site !== "") {
        map[ord].cidadesMap[cidadeVal].sites.add(l.site);
      }
      if (typeof l.mttr === "number") {
        map[ord].cidadesMap[cidadeVal].mttrValores.push(l.mttr);
      }
    });

    function porCidadeDoMes(cidadesMap: { [cidade: string]: { dentro: number; fora: number; sites: Set<string>; mttrValores: number[]; } }) {
      return Object.keys(cidadesMap).map(cidade => {
        const c = cidadesMap[cidade];
        const total = c.dentro + c.fora;
        const mttrMedioSeg = c.mttrValores.length > 0
          ? c.mttrValores.reduce((a, b) => a + b, 0) / c.mttrValores.length
          : null;
        return {
          cidade: cidade,
          qtd: total,
          sites: c.sites.size,
          mttr: mttrMedioSeg !== null ? mttrMedioSeg / 3600 : null,
          eser4h: total > 0 ? c.dentro / total : 0
        };
      }).sort((a, b) => b.qtd - a.qtd);
    }

    return Object.keys(map).sort().map(k => ({
      mes: map[k].mes,
      areas: map[k].areas,
      diarioAreas: Object.keys(map[k].diarioAreasMap).sort().map(dia => ({ dia: dia, areas: map[k].diarioAreasMap[dia] })),
      mttrValores: mttrPorMes[k] || [],
      horasSubmit: horasSubmitStatsPorMes[k] || [],
      horasClear: horasClearStatsPorMes[k] || [],
      topSites: topSitesDoMes(k),
      todosSites: todosSitesDoMes(k), // NOVO (item 26)
      totalSitesUnicos: totalSitesUnicosDoMes(k),
      genericTier2Top10: topContagemDoMes(genericTier2PorMes, k),
      genericTier3Top10: topContagemDoMes(genericTier3PorMes, k),
      resolucaoTop10: topContagemDoMes(resolucaoPorMes, k),
      porCidade: porCidadeDoMes(map[k].cidadesMap) // NOVO (item 18)
    }));
  }

  const meses = buildMeses("eser4h");
  const meses6h = buildMeses("eser6h");

  // ALTERADO (item 43): a semana é sempre a ISO da dataRef (CLEAR_DATE,
  // senão RESOLVED_DATE). Incidente com clear em outubro cai na semana de outubro.
  function buildSemanal(campo: "eser4h" | "eser6h") {
    const map: { [key: string]: { semana: string; dentro: number; fora: number; ultimaData: number; } } = {};
    linhas.forEach(l => {
      const val = l[campo];
      if (val !== "Dentro" && val !== "Fora") {
        return;
      }
      const wk = getISOWeek(l.dataRef);
      const key = "Semana " + wk;
      if (!map[key]) {
        map[key] = { semana: key, dentro: 0, fora: 0, ultimaData: 0 };
      }
      if (val === "Dentro") {
        map[key].dentro++;
      } else {
        map[key].fora++;
      }
      const t = l.dataRef.getTime();
      if (t > map[key].ultimaData) {
        map[key].ultimaData = t;
      }
    });
    // ordena pela data (e não pelo número) pra não embaralhar na virada do ano
    const todas = Object.keys(map).map(k => map[k]).sort((a, b) => a.ultimaData - b.ultimaData);
    const ultimas8 = todas.slice(-8);
    return ultimas8.map(x => ({ semana: x.semana, dentro: x.dentro, fora: x.fora }));
  }

  const semanal = buildSemanal("eser4h");
  const semanal6h = buildSemanal("eser6h");

  const hoje = new Date();
  const atualizadoEm = hoje.getFullYear() + "-" + pad2(hoje.getMonth() + 1) + "-" + pad2(hoje.getDate());

  // NOVO (item 27): data do incidente mais recente da base (maior dataRef).
  let maxTempoDataRef = 0;
  for (let i = 0; i < linhas.length; i++) {
    const t = linhas[i].dataRef.getTime();
    if (t > maxTempoDataRef) {
      maxTempoDataRef = t;
    }
  }
  const dMaxDataRef = new Date(maxTempoDataRef);
  const dadosAte = maxTempoDataRef > 0
    ? pad2(dMaxDataRef.getDate()) + "/" + pad2(dMaxDataRef.getMonth() + 1) + "/" + dMaxDataRef.getFullYear()
    : "";

  // ALTERADO (item 43): a coluna "Semana" da Historico não é mais usada.
  const semanaFonte = "ISO calculada da CLEAR_DATE (ou RESOLVED_DATE quando vazia)";

  // ALTERADO (item 43): "sitesBase" — contagem por SITE_CI direto da aba
  // "Historico" (Tipo = Indisponibilidade, sem filtrar Área, sem
  // EXCLUIR_INCIDENTES). O mês vem da CLEAR_DATE ou, se vazia, da
  // RESOLVED_DATE. Sem nenhuma das duas, a linha vai pra "sitesBaseSemData".
  const sitesBasePorMes: { [ord: string]: { [site: string]: number } } = {};
  const sitesBaseSemData: { [site: string]: number } = {};
  for (let r = 0; r < rowCount - 1; r++) {
    if (tipoEfetivo(colTipo[r], colIncidente[r]) !== "Indisponibilidade") { // ALTERADO (item 46)
      continue;
    }
    const siteBase = String(colSite[r] || "").trim();
    if (siteBase === "") {
      continue;
    }
    const incBase = normaliza(String(colIncidente[r] || "").trim());
    if (incBase !== "" && incidentesExcluidos.has(incBase)) {
      continue;
    }
    const datasCandidatas = [colClear[r], colResolved[r]];
    let serialBase: number | null = null;
    for (let k = 0; k < datasCandidatas.length; k++) {
      const v = datasCandidatas[k];
      if (typeof v === "number" && v > 0) {
        serialBase = v;
        break;
      }
    }
    if (serialBase === null) {
      sitesBaseSemData[siteBase] = (sitesBaseSemData[siteBase] || 0) + 1;
      continue;
    }
    let dataBase = serialToDate(serialBase);
    if (incidentesAgosto.has(incBase)) {
      dataBase = new Date(2026, 7, 1); // regra 2: esses incidentes contam em Agosto
    }
    if (incidentesOutubro.has(incBase)) {
      dataBase = ajustarParaOutubro(dataBase); // NOVO (item 45): contam em Outubro
    }
    if (incidentesNovembro.has(incBase)) {
      dataBase = ajustarParaNovembro(dataBase); // NOVO (item 49): contam em Novembro
    }
    const ordBase = dataBase.getFullYear() + "-" + pad2(dataBase.getMonth() + 1);
    if (!sitesBasePorMes[ordBase]) {
      sitesBasePorMes[ordBase] = {};
    }
    sitesBasePorMes[ordBase][siteBase] = (sitesBasePorMes[ordBase][siteBase] || 0) + 1;
  }
  const sitesBase = Object.keys(sitesBasePorMes).sort().map(ord => ({
    mesOrd: ord,
    mes: MES_PT[Number(ord.substring(5, 7)) - 1] + "/" + ord.substring(2, 4),
    sites: Object.keys(sitesBasePorMes[ord])
      .map(site => ({ site: site, total: sitesBasePorMes[ord][site] }))
      .sort((a, b) => b.total - a.total)
  }));

  const ABAS_ORGANOGRAMA = ["Organograma", "PJ"];
  const colaboradores: {
    contrato: string;
    areaTecnica: string;
    supervisor: string;
    unidadeNegocio: string;
    coordenador: string;
    dataAdmis: string;
    dataDemissao: string;
    ativo: boolean;
    matricula: string;
    nome: string;
    placa: string;
    descFuncao: string;
    endereco: string; // NOVO (item 31)
    escala: string;   // NOVO (TOA aderência)
    turno: string;    // NOVO (TOA aderência)
    horario: string;  // NOVO (TOA aderência)
  }[] = [];

  function serialParaIso(serial: unknown): string {
    if (typeof serial !== "number" || serial <= 0) {
      return "";
    }
    const d = serialToDate(serial);
    return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
  }

  function processarAbaColaboradores(nomeAba: string): void {
    let sheetOrganograma = workbook.getWorksheet(nomeAba);
    if (!sheetOrganograma) {
      const alvoOrg = normaliza(nomeAba);
      for (const ws of todasAbas) {
        if (normaliza(ws.getName()) === alvoOrg) {
          sheetOrganograma = ws;
          break;
        }
      }
    }
    if (!sheetOrganograma) {
      return;
    }
    const rangeOrg = sheetOrganograma.getUsedRange();
    if (!rangeOrg) {
      return;
    }
    const rowCountOrg = rangeOrg.getRowCount();
    const colCountOrg = rangeOrg.getColumnCount();
    if (rowCountOrg < 2) {
      return;
    }
    const headersOrg = sheetOrganograma.getRangeByIndexes(0, 0, 1, colCountOrg).getValues()[0];
    const colIndexOrg: { [key: string]: number } = {};
    headersOrg.forEach((h, i) => {
      const chave = normaliza(h);
      if (chave !== "") {
        colIndexOrg[chave] = i;
      }
    });
    function getIdxOrg(nomeColuna: string): number | undefined {
      return colIndexOrg[normaliza(nomeColuna)];
    }
    const idxDataAdmis = getIdxOrg("Data Admis.");
    const idxDataDemissao = getIdxOrg("Dt. Demissao");
    const idxStatus = getIdxOrg("Status");
    const idxAreaTecnica = getIdxOrg("Área Técnica");
    const idxSupervisorOrg = getIdxOrg("Supervisor");
    const idxContrato = getIdxOrg("Contrato_2");
    const idxUnidadeNegocio = getIdxOrg("Unidade de Negócio");
    const idxCoordenador = getIdxOrg("Coordenador");
    const idxMatricula = getIdxOrg("Matricula");
    const idxNome = getIdxOrg("Nome");
    const idxPlaca = getIdxOrg("Placa");
    const idxDescFuncao = getIdxOrg("Desc.Funcao");
    const idxEndereco = getIdxOrg("Endereço"); // NOVO (item 31)
    const idxNumEndereco = getIdxOrg("Num.Endereço"); // NOVO (item 31)
    const idxEscalaOrg = getIdxOrg("Escala");
    const idxTurnoOrg = getIdxOrg("Turno");
    const idxHorarioOrg = getIdxOrg("Horario") !== undefined ? getIdxOrg("Horario") : getIdxOrg("Horário");

    function getColunaOrg(idx: number | undefined): (string | number)[] {
      if (idx === undefined) {
        return [];
      }
      return sheetOrganograma!.getRangeByIndexes(1, idx, rowCountOrg - 1, 1).getValues().map(r => r[0]) as (string | number)[];
    }

    const colDataAdmis = getColunaOrg(idxDataAdmis);
    const colDataDemissao = getColunaOrg(idxDataDemissao);
    const colStatus = getColunaOrg(idxStatus);
    const colAreaTecnica = getColunaOrg(idxAreaTecnica);
    const colSupervisorOrg = getColunaOrg(idxSupervisorOrg);
    const colContrato = getColunaOrg(idxContrato);
    const colUnidadeNegocio = getColunaOrg(idxUnidadeNegocio);
    const colCoordenador = getColunaOrg(idxCoordenador);
    const colMatricula = getColunaOrg(idxMatricula);
    const colNome = getColunaOrg(idxNome);
    const colPlaca = getColunaOrg(idxPlaca);
    const colDescFuncao = getColunaOrg(idxDescFuncao);
    const colEndereco = getColunaOrg(idxEndereco);
    const colNumEndereco = getColunaOrg(idxNumEndereco);
    const colEscalaOrg = getColunaOrg(idxEscalaOrg);
    const colTurnoOrg = getColunaOrg(idxTurnoOrg);
    const colHorarioOrg = getColunaOrg(idxHorarioOrg);

    function horarioTexto(v: string | number): string {
      // horário pode vir como fração do dia do Excel (0,3333 = 08:00) ou como texto ("08:00 AS 17:48")
      if (typeof v === "number" && v > 0 && v < 1) {
        const min = Math.round(v * 24 * 60);
        return pad2(Math.floor(min / 60)) + ":" + pad2(min % 60);
      }
      return String(v || "").trim();
    }
    for (let r = 0; r < rowCountOrg - 1; r++) {
      const statusVal = String(colStatus[r] || "").trim();
      const ativo = normaliza(statusVal) === "ativo";
      const contratoVal = String(colContrato[r] || "").trim();
      const areaTecnicaVal = String(colAreaTecnica[r] || "").trim();
      const supervisorVal = String(colSupervisorOrg[r] || "").trim();
      const unidadeNegocioVal = String(colUnidadeNegocio[r] || "").trim();
      const coordenadorVal = String(colCoordenador[r] || "").trim();
      const matriculaVal = String(colMatricula[r] || "").trim();
      const nomeVal = String(colNome[r] || "").trim();
      const placaVal = String(colPlaca[r] || "").trim();
      const descFuncaoVal = String(colDescFuncao[r] || "").trim();
      const enderecoRua = String(colEndereco[r] || "").trim();
      const enderecoNum = String(colNumEndereco[r] || "").trim();
      const enderecoVal = enderecoRua !== "" && enderecoNum !== "" ? enderecoRua + ", " + enderecoNum : (enderecoRua !== "" ? enderecoRua : enderecoNum);

      if (contratoVal === "" && areaTecnicaVal === "" && supervisorVal === "") {
        continue;
      }

      colaboradores.push({
        contrato: contratoVal,
        areaTecnica: areaTecnicaVal,
        supervisor: supervisorVal,
        unidadeNegocio: unidadeNegocioVal,
        coordenador: coordenadorVal,
        dataAdmis: serialParaIso(colDataAdmis[r]),
        dataDemissao: serialParaIso(colDataDemissao[r]),
        ativo: ativo,
        matricula: matriculaVal,
        nome: nomeVal,
        placa: placaVal,
        descFuncao: descFuncaoVal,
        endereco: enderecoVal,
        escala: String(colEscalaOrg[r] || "").trim(),
        turno: String(colTurnoOrg[r] || "").trim(),
        horario: horarioTexto(colHorarioOrg[r])
      });
    }
  }

  ABAS_ORGANOGRAMA.forEach(nomeAba => {
    processarAbaColaboradores(nomeAba);
  });

  interface Veiculo {
    placa: string;
    condutor: string;
    status: string;
    marca: string;
    modelo: string;
    combustivel: string;
    tipo: string;
    operacao: string;
    valorCombustivel: number | null;
    valor: number | string;
  }

  const frota: Veiculo[] = [];

  function processarAbaFrota(): void {
    const NOME_ABA_FROTA = "Placa";
    let sheetFrota = workbook.getWorksheet(NOME_ABA_FROTA);
    if (!sheetFrota) {
      const alvoFrota = normaliza(NOME_ABA_FROTA);
      for (const ws of todasAbas) {
        if (normaliza(ws.getName()) === alvoFrota) {
          sheetFrota = ws;
          break;
        }
      }
    }
    if (!sheetFrota) {
      return;
    }
    const rangeFrota = sheetFrota.getUsedRange();
    if (!rangeFrota) {
      return;
    }
    const rowCountFrota = rangeFrota.getRowCount();
    const colCountFrota = rangeFrota.getColumnCount();
    if (rowCountFrota < 2) {
      return;
    }
    const headersFrota = sheetFrota.getRangeByIndexes(0, 0, 1, colCountFrota).getValues()[0];
    const colIndexFrota: { [key: string]: number } = {};
    headersFrota.forEach((h, i) => {
      const chave = normaliza(h);
      if (chave !== "") {
        colIndexFrota[chave] = i;
      }
    });
    function getIdxFrota(nomeColuna: string): number | undefined {
      return colIndexFrota[normaliza(nomeColuna)];
    }
    const idxPlacaFrota = getIdxFrota("Placa");
    const idxCondutor = getIdxFrota("Condutor do veículo");
    const idxStatusFrota = getIdxFrota("Status");
    const idxMarca = getIdxFrota("Marca");
    const idxModelo = getIdxFrota("Modelo");
    const idxCombustivel = getIdxFrota("Combustivel");
    const idxTipoFrota = getIdxFrota("Tipo");
    const idxOperacao = getIdxFrota("Operação");
    const idxValor = getIdxFrota("Valor da locação ou valor de compra (se próprio)");

    function getColunaFrota(idx: number | undefined): (string | number)[] {
      if (idx === undefined) {
        return [];
      }
      return sheetFrota!.getRangeByIndexes(1, idx, rowCountFrota - 1, 1).getValues().map(r => r[0]) as (string | number)[];
    }

    const colPlacaFrota = getColunaFrota(idxPlacaFrota);
    const colCondutor = getColunaFrota(idxCondutor);
    const colStatusFrota = getColunaFrota(idxStatusFrota);
    const colMarca = getColunaFrota(idxMarca);
    const colModelo = getColunaFrota(idxModelo);
    const colCombustivel = getColunaFrota(idxCombustivel);
    const colTipoFrota = getColunaFrota(idxTipoFrota);
    const colOperacao = getColunaFrota(idxOperacao);
    const colValor = getColunaFrota(idxValor);

    for (let r = 0; r < rowCountFrota - 1; r++) {
      const placaVal = String(colPlacaFrota[r] || "").trim();
      const condutorVal = String(colCondutor[r] || "").trim();
      const statusVal = String(colStatusFrota[r] || "").trim();
      const marcaVal = String(colMarca[r] || "").trim();
      const modeloVal = String(colModelo[r] || "").trim();
      const combustivelVal = String(colCombustivel[r] || "").trim();
      const tipoVal = String(colTipoFrota[r] || "").trim();
      const operacaoVal = String(colOperacao[r] || "").trim();

      const valorBruto = colValor[r];
      const valorVal: number | string = typeof valorBruto === "number" ? valorBruto : String(valorBruto || "").trim();

      if (placaVal === "" && modeloVal === "") {
        continue;
      }

      frota.push({
        placa: placaVal,
        condutor: condutorVal,
        status: statusVal,
        marca: marcaVal,
        modelo: modeloVal,
        combustivel: combustivelVal,
        tipo: tipoVal,
        operacao: operacaoVal,
        valorCombustivel: rateioPorPlaca[normaliza(placaVal)] !== undefined ? rateioPorPlaca[normaliza(placaVal)] : null,
        valor: valorVal
      });
    }
  }

  processarAbaFrota();

  interface Vaga {
    recrutadora: string;
    dataCriacaoRequisicao: string;
    requisicaoAprovada: string;
    previsaoEntrega: string;
    requisicao: string;
    area: string;
    supervisao: string;
    cargo: string;
    numeroPosicoes: string | number;
    motivo: string;
    profissionalSubstituido: string;
    modelo: string;
    column12: string;
    column13: string;
    statusReq: string;
    fechamento: string;
    status: string;
    nomeCandidatoAprovado: string;
    cpf: string;
    slaDias: string | number;
    envioLinkDocumento: string;
    statusAso: string;
    dataExame: string;
    previsaoAdmissao: string;
    slaAdmissao: string | number;
    dataIntegracao: string;
    liberado: string;
    slaTotal: string | number;
    column28: string;
    column29: string;
  }

  const vagas: Vaga[] = [];

  function processarAbaVagas(): void {
    const NOME_ABA_VAGAS = "Controle";
    let sheetVagas = workbook.getWorksheet(NOME_ABA_VAGAS);
    if (!sheetVagas) {
      const alvoVagas = normaliza(NOME_ABA_VAGAS);
      for (const ws of todasAbas) {
        if (normaliza(ws.getName()) === alvoVagas) {
          sheetVagas = ws;
          break;
        }
      }
    }
    if (!sheetVagas) {
      return;
    }
    const rangeVagas = sheetVagas.getUsedRange();
    if (!rangeVagas) {
      return;
    }
    const rowCountVagas = rangeVagas.getRowCount();
    const colCountVagas = rangeVagas.getColumnCount();
    if (rowCountVagas < 2) {
      return;
    }
    const headersVagas = sheetVagas.getRangeByIndexes(0, 0, 1, colCountVagas).getValues()[0];
    const colIndexVagas: { [key: string]: number } = {};
    headersVagas.forEach((h, i) => {
      const chave = normaliza(h);
      if (chave !== "") {
        colIndexVagas[chave] = i;
      }
    });
    function getIdxVagas(...nomes: string[]): number | undefined {
      // aceita mais de um nome de cabeçalho (a planilha "Controle" mudou os títulos em 11/2026)
      for (const nome of nomes) {
        const i = colIndexVagas[normaliza(nome)];
        if (i !== undefined) return i;
      }
      return undefined;
    }
    const idxRecrutadora = getIdxVagas("Recrutadora");
    const idxDataCriacaoRequisicao = getIdxVagas("Data de criação da Requisição", "Data de criação da Req.", "Data de criação da Req", "Data de criação");
    const idxRequisicaoAprovada = getIdxVagas("Requisição aprovada");
    const idxPrevisaoEntrega = getIdxVagas("Previsão de entrega");
    const idxRequisicao = getIdxVagas("Requisição");
    const idxArea = getIdxVagas("ÁREA");
    const idxSupervisao = getIdxVagas("SUPERVISÃO", "Supervisor");
    const idxCargo = getIdxVagas("Cargo (Obrigatório)", "Cargo");
    const idxNumeroPosicoes = getIdxVagas("Número de posições");
    const idxMotivo = getIdxVagas("Motivo");
    const idxProfissionalSubstituido = getIdxVagas("PROFISSIONAL SUBSTITUÍDO", "Profissional substituido");
    const idxModelo = getIdxVagas("Modelo");
    const idxColumn12 = getIdxVagas("Column12");
    const idxColumn13 = getIdxVagas("Column13");
    const idxStatusReq = getIdxVagas("Status Req.");
    const idxFechamento = getIdxVagas("Fechamento", "Data de Fechamento", "Data Fechamento");
    const idxStatus = getIdxVagas("STATUS", "Status no gestão de vagas", "Status Vaga");
    const idxNomeCandidatoAprovado = getIdxVagas("Nome do candidato aprovado");
    const idxCpf = getIdxVagas("CPF");
    const idxSlaDias = getIdxVagas("SLA dias");
    const idxEnvioLinkDocumento = getIdxVagas("Envio de link documento");
    const idxStatusAso = getIdxVagas("Status ASO");
    const idxDataExame = getIdxVagas("Data exame");
    const idxPrevisaoAdmissao = getIdxVagas("Previsão de admissão");
    const idxSlaAdmissao = getIdxVagas("SLA Admissão");
    const idxDataIntegracao = getIdxVagas("Data Integração");
    const idxLiberado = getIdxVagas("Liberado?");
    const idxSlaTotal = getIdxVagas("SLA total");
    const idxColumn28 = getIdxVagas("Column28");
    const idxColumn29 = getIdxVagas("Column29");

    function getColunaVagas(idx: number | undefined): (string | number)[] {
      if (idx === undefined) {
        return [];
      }
      return sheetVagas!.getRangeByIndexes(1, idx, rowCountVagas - 1, 1).getValues().map(r => r[0]) as (string | number)[];
    }

    function formatarDataVaga(valorBruto: string | number): string {
      if (typeof valorBruto === "number" && valorBruto > 0) {
        const d = serialToDate(valorBruto);
        return pad2(d.getDate()) + "/" + pad2(d.getMonth() + 1) + "/" + d.getFullYear();
      }
      return String(valorBruto || "").trim();
    }

    function getCampoGenerico(valorBruto: string | number): string | number {
      if (typeof valorBruto === "number") {
        return valorBruto;
      }
      return String(valorBruto || "").trim();
    }

    const colRecrutadora = getColunaVagas(idxRecrutadora);
    const colDataCriacaoRequisicao = getColunaVagas(idxDataCriacaoRequisicao);
    const colRequisicaoAprovada = getColunaVagas(idxRequisicaoAprovada);
    const colPrevisaoEntrega = getColunaVagas(idxPrevisaoEntrega);
    const colRequisicao = getColunaVagas(idxRequisicao);
    const colArea = getColunaVagas(idxArea);
    const colSupervisao = getColunaVagas(idxSupervisao);
    const colCargo = getColunaVagas(idxCargo);
    const colNumeroPosicoes = getColunaVagas(idxNumeroPosicoes);
    const colMotivo = getColunaVagas(idxMotivo);
    const colProfissionalSubstituido = getColunaVagas(idxProfissionalSubstituido);
    const colModelo = getColunaVagas(idxModelo);
    const colColumn12 = getColunaVagas(idxColumn12);
    const colColumn13 = getColunaVagas(idxColumn13);
    const colStatusReq = getColunaVagas(idxStatusReq);
    const colFechamento = getColunaVagas(idxFechamento);
    const colStatus = getColunaVagas(idxStatus);
    const colNomeCandidatoAprovado = getColunaVagas(idxNomeCandidatoAprovado);
    const colCpf = getColunaVagas(idxCpf);
    const colSlaDias = getColunaVagas(idxSlaDias);
    const colEnvioLinkDocumento = getColunaVagas(idxEnvioLinkDocumento);
    const colStatusAso = getColunaVagas(idxStatusAso);
    const colDataExame = getColunaVagas(idxDataExame);
    const colPrevisaoAdmissao = getColunaVagas(idxPrevisaoAdmissao);
    const colSlaAdmissao = getColunaVagas(idxSlaAdmissao);
    const colDataIntegracao = getColunaVagas(idxDataIntegracao);
    const colLiberado = getColunaVagas(idxLiberado);
    const colSlaTotal = getColunaVagas(idxSlaTotal);
    const colColumn28 = getColunaVagas(idxColumn28);
    const colColumn29 = getColunaVagas(idxColumn29);

    for (let r = 0; r < rowCountVagas - 1; r++) {
      const requisicaoVal = String(colRequisicao[r] || "").trim();
      const cargoVal = String(colCargo[r] || "").trim();

      if (requisicaoVal === "" && cargoVal === "") {
        continue;
      }

      vagas.push({
        recrutadora: String(colRecrutadora[r] || "").trim(),
        dataCriacaoRequisicao: formatarDataVaga(colDataCriacaoRequisicao[r]),
        requisicaoAprovada: formatarDataVaga(colRequisicaoAprovada[r]),
        previsaoEntrega: formatarDataVaga(colPrevisaoEntrega[r]),
        requisicao: requisicaoVal,
        area: String(colArea[r] || "").trim(),
        supervisao: String(colSupervisao[r] || "").trim(),
        cargo: cargoVal,
        numeroPosicoes: getCampoGenerico(colNumeroPosicoes[r]),
        motivo: String(colMotivo[r] || "").trim(),
        profissionalSubstituido: String(colProfissionalSubstituido[r] || "").trim(),
        modelo: String(colModelo[r] || "").trim(),
        column12: String(colColumn12[r] || "").trim(),
        column13: String(colColumn13[r] || "").trim(),
        statusReq: String(colStatusReq[r] || "").trim(),
        fechamento: formatarDataVaga(colFechamento[r]),
        status: String(colStatus[r] || "").trim(),
        nomeCandidatoAprovado: String(colNomeCandidatoAprovado[r] || "").trim(),
        cpf: String(colCpf[r] || "").trim(),
        slaDias: getCampoGenerico(colSlaDias[r]),
        envioLinkDocumento: String(colEnvioLinkDocumento[r] || "").trim(),
        statusAso: String(colStatusAso[r] || "").trim(),
        dataExame: formatarDataVaga(colDataExame[r]),
        previsaoAdmissao: formatarDataVaga(colPrevisaoAdmissao[r]),
        slaAdmissao: getCampoGenerico(colSlaAdmissao[r]),
        dataIntegracao: formatarDataVaga(colDataIntegracao[r]),
        liberado: String(colLiberado[r] || "").trim(),
        slaTotal: getCampoGenerico(colSlaTotal[r]),
        column28: String(colColumn28[r] || "").trim(),
        column29: String(colColumn29[r] || "").trim()
      });
    }
  }

  processarAbaVagas();

  let bancoHoras: unknown[] = [];
  try {
    bancoHoras = JSON.parse(bancoHorasJson || "[]") as unknown[];
  } catch (e) {
    bancoHoras = [];
  }

  // NOVO (item 41): Análise de Turnover.
  let turnover: unknown[] = [];
  try {
    const turnoverBruto = JSON.parse(turnoverJson || "[]") as unknown;
    turnover = Array.isArray(turnoverBruto) ? turnoverBruto : [];
  } catch (e) {
    turnover = [];
  }

  // ==================================================================
  // NOVO (item 48): aba "TOA" — atividades lançadas pelos técnicos.
  // Ignora "Refeição" (Tipo de Atividade). Agrupa por Data + Recurso:
  // "total" = atividades lançadas (qualquer status) e
  // "concluidas" = Status "Concluída".
  // Supervisor = última coluna "Supervisor" da aba.
  // ==================================================================
  interface ToaItem {
    data: string;
    recurso: string;
    supervisor: string;
    total: number;
    concluidas: number;
    refeicao: number;
    status: { [status: string]: number };
  }
  const toaPorChave: { [chave: string]: ToaItem } = {};
  const debugToa: {
    abaEncontrada: boolean;
    colunasFaltando: string[];
    linhasLidas: number;
    linhasUsadas: number;
    linhasSemData: number;
  } = {
    abaEncontrada: false,
    colunasFaltando: [],
    linhasLidas: 0,
    linhasUsadas: 0,
    linhasSemData: 0
  };

  function dataToaIso(valor: unknown): string {
    if (typeof valor === "number" && valor > 0) {
      return serialParaIso(valor);
    }
    const txt = String(valor ?? "").trim();
    const br = txt.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (br) {
      return br[3] + "-" + pad2(Number(br[2])) + "-" + pad2(Number(br[1]));
    }
    const iso = txt.match(/^(\d{4})-(\d{2})-(\d{2})/);
    return iso ? iso[1] + "-" + iso[2] + "-" + iso[3] : "";
  }

  function processarAbaTOA(): void {
    let sheetToa = workbook.getWorksheet("TOA");
    if (!sheetToa) {
      for (const ws of todasAbas) {
        if (normaliza(ws.getName()) === "toa") {
          sheetToa = ws;
          break;
        }
      }
    }
    if (!sheetToa) {
      return;
    }
    debugToa.abaEncontrada = true;
    const rangeToa = sheetToa.getUsedRange();
    if (!rangeToa || rangeToa.getRowCount() < 2) {
      return;
    }
    const nLin = rangeToa.getRowCount();
    const nCol = rangeToa.getColumnCount();
    const valoresToa = sheetToa.getRangeByIndexes(0, 0, nLin, nCol).getValues();
    const cabToa = valoresToa[0].map(h => normaliza(h));

    function idxToa(nomes: string[]): number {
      for (const n of nomes) {
        const i = cabToa.indexOf(normaliza(n));
        if (i !== -1) {
          return i;
        }
      }
      return -1;
    }

    const iTipo = idxToa(["Tipo de Atividade", "Tipo Atividade"]);
    const iStatus = idxToa(["Status"]);
    const iData = idxToa(["Data", "Data dia", "Dia"]);
    const iRecurso = idxToa(["Recurso", "Recurso (funcionário)", "Funcionário", "Técnico"]);
    const iSup = cabToa.lastIndexOf("supervisor");

    if (iTipo === -1) { debugToa.colunasFaltando.push("Tipo de Atividade"); }
    if (iStatus === -1) { debugToa.colunasFaltando.push("Status"); }
    if (iData === -1) { debugToa.colunasFaltando.push("Data"); }
    if (iRecurso === -1) { debugToa.colunasFaltando.push("Recurso"); }
    if (iSup === -1) { debugToa.colunasFaltando.push("Supervisor"); }
    if (iData === -1 || iRecurso === -1) {
      return;
    }

    for (let r = 1; r < nLin; r++) {
      const linha = valoresToa[r];
      debugToa.linhasLidas++;
      const tipoAtividade = iTipo !== -1 ? normaliza(linha[iTipo]) : "";
      const ehRefeicao = tipoAtividade.indexOf("refeicao") !== -1;
      const recurso = String(linha[iRecurso] ?? "").trim();
      if (recurso === "") {
        continue;
      }
      const dataIso = dataToaIso(linha[iData]);
      if (dataIso === "") {
        debugToa.linhasSemData++;
        continue;
      }
      const supervisorToa = iSup !== -1 ? String(linha[iSup] ?? "").trim() : "";
      const chave = dataIso + "|" + normaliza(recurso);
      if (!toaPorChave[chave]) {
        toaPorChave[chave] = { data: dataIso, recurso: recurso, supervisor: supervisorToa, total: 0, concluidas: 0, refeicao: 0, status: {} };
      }
      const item = toaPorChave[chave];
      if (item.supervisor === "" && supervisorToa !== "") {
        item.supervisor = supervisorToa;
      }
      if (ehRefeicao) {
        item.refeicao++;
      } else {
        item.total++;
        const statusTxt = iStatus !== -1 ? String(linha[iStatus] ?? "").trim() : "";
        const statusChave = statusTxt !== "" ? statusTxt : "Sem status";
        item.status[statusChave] = (item.status[statusChave] || 0) + 1;
        if (normaliza(statusTxt).indexOf("conclu") === 0) {
          item.concluidas++;
        }
      }
      debugToa.linhasUsadas++;
    }
  }

  processarAbaTOA();
  const toa = Object.keys(toaPorChave).sort().map(k => toaPorChave[k]);

  // NOVO (temporário): diagnóstico do cruzamento Fora do Prazo x aba Fila.
  const debugMatchFila: {
    totalForaDoPrazo: number;
    totalComMatchNaFila: number;
    exemploComMatch: string[];
    exemploSemMatch: string[];
  } = {
    totalForaDoPrazo: pendentesFca.length,
    totalComMatchNaFila: 0,
    exemploComMatch: [],
    exemploSemMatch: []
  };
  pendentesFca.forEach(p => {
    const chave = normaliza(p.incidente);
    const bateu = filaPorIncidente[chave] !== undefined;
    if (bateu) {
      debugMatchFila.totalComMatchNaFila++;
      if (debugMatchFila.exemploComMatch.length < 5) {
        debugMatchFila.exemploComMatch.push(p.incidente);
      }
    } else if (debugMatchFila.exemploSemMatch.length < 5) {
      debugMatchFila.exemploSemMatch.push(p.incidente);
    }
  });

  const PRES_TOA = lerPresencaTOACompleta(workbook);
  const resultado = {
    meta: META,
    meses: meses,
    diario: diario,
    diario6h: diario6h,
    semanal: semanal,
    meses6h: meses6h,
    semanal6h: semanal6h,
    pendentesFca: pendentesFca,
    incidentesPorSite: incidentesPorSite,
    fcaTipoDivergente: fcaTipoDivergente,
    colaboradores: colaboradores,
    frota: frota,
    vagas: vagas,
    bancoHoras: bancoHoras,
    turnover: turnover,
    toa: toa, // NOVO (item 48)
    quedasSite: quedasSite, // NOVO (11/11/2026) – Acompanhamento de Alarmes (Indisponibilidade + Outros)
    toaPresenca: PRES_TOA.logaram, // técnicos com bonequinho no TOA (aba TOA_PRESENCA) – azul, laranja ou preto
    presencaToaBase: PRES_TOA.logaram, // cópia da mesma lista (o painel usa se toaPresenca vier vazio)
    toaNaoIniciou: PRES_TOA.naoIniciou, // ícone "ptnon_working_cal" (rota não iniciada)
    debugToaPresenca: PRES_TOA.debug, // diagnóstico mostrado no painel
    avaliacoesExperiencia: lerAvaliacoesExperiencia(workbook), // NOVO – avaliação 30/60/90 dias (aba AVALIACAO_EXPERIENCIA)
    entrevistasDesligamento: lerAvaliacoesExperiencia(workbook, "SAIDA_TECNICO"), // NOVO – saída do técnico (aba SAIDA_TECNICO)
    atualizado_em: atualizadoEm,
    dadosAte: dadosAte,
    semanaFonte: semanaFonte,
    debugFcaMatch: debugFcaMatch, // TEMPORÁRIO (item 35)
    debugFcaCampos: { chavesRecebidas: fcaChavesRecebidas, itensComCausa: fcaItensComCausa }, // TEMPORÁRIO (item 36)
    debugEscritaFca: debugEscritaFca, // TEMPORÁRIO (item 37)
    debugToa: debugToa, // TEMPORÁRIO (item 48)
    sitesBase: sitesBase,
    sitesBaseSemData: sitesBaseSemData,
    debugFila: debugFila, // TEMPORÁRIO
    debugMatchFila: debugMatchFila // TEMPORÁRIO
  };

  return JSON.stringify(resultado);
}

// ===================================================================
// TOA – PRESENÇA (bonequinho) → dados.json "toaPresenca"
// Cole esta função no FINAL do Office Script que gera o dados.json (item 48)
// e, no objeto que vira o dados.json (onde já tem  toa: ...),
// acrescente a linha:      toaPresenca: lerPresencaTOA(workbook),
//
// Aba "TOA_PRESENCA" (gravada pelo robô), cabeçalho na linha 1:
//   Data | Tecnico | Icone
// Conta como aderente o técnico com bonequinho (class "res-ic ...").
// ===================================================================
function lerPresencaTOACompleta(workbook: ExcelScript.Workbook): {
  logaram: { data: string; recurso: string }[];
  naoIniciou: { data: string; recurso: string }[];
  debug: { aba: string; linhas: number; logaram: number; naoLogaram: number; naoIniciaram: number; semData: number; aviso: string };
} {
  const vazio = (aviso: string, aba: string) => ({
    logaram: [] as { data: string; recurso: string }[], naoIniciou: [] as { data: string; recurso: string }[],
    debug: { aba: aba, linhas: 0, logaram: 0, naoLogaram: 0, naoIniciaram: 0, semData: 0, aviso: aviso }
  });
  let ws = workbook.getWorksheet("TOA_PRESENCA");
  if (!ws) {   // aceita a aba com outro jeito de escrever (ex.: "TOA PRESENÇA")
    const achada = workbook.getWorksheets().filter(w => /TOA.?PRESEN/.test(w.getName().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase()));
    if (achada.length) ws = achada[0];
  }
  if (!ws) return vazio("aba TOA_PRESENCA não encontrada", "");
  const nomeAba = ws.getName();
  const usado = ws.getUsedRange();
  if (!usado) return vazio("aba vazia", nomeAba);
  const v = usado.getValues();
  if (v.length < 1) return vazio("aba vazia", nomeAba);
  const sa = (x: string | number | boolean) => String(x == null ? "" : x).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();
  const cab = v[0].map(x => sa(x));
  let cData = cab.findIndex(c => c.indexOf("DATA") === 0);
  let cTec = cab.findIndex(c => /TECNICO|RECURSO|NOME/.test(c));
  let cIco = cab.findIndex(c => /ICONE|STATUS|COR/.test(c));
  let ini = 1;
  let aviso = "";
  if (cData < 0 || cTec < 0) {   // sem cabeçalho: usa A = Data, B = Técnico, C = Ícone
    cData = 0; cTec = 1; cIco = v[0].length > 2 ? 2 : -1; ini = 0;
    aviso = "cabeçalho não encontrado – usei A=Data, B=Técnico, C=Ícone";
  }
  // Ícone do TOA (atributo class):
  //   bonequinho AZUL  = "res-ic res_type_10 pt1_32_..."  (rota iniciada)   → LOGOU
  //   bonequinho PRETO = "res-ic res_type_10 pt1_128_..." (rota encerrada)  → LOGOU
  //   bonequinho LARANJA = "res-ic res_type_10 pt1_64_..." (rota NÃO iniciada) → toaNaoIniciou
  //   quadradinho "ptnon_working_cal" / X "ptnon_configured_cal" → não logou
  const LOGOU = /RES-IC|RES_TYPE_/;
  const NAO_INICIOU = /PT1_64_/;
  const iso = (x: string | number | boolean): string => {
    if (typeof x === "number" && x > 20000) {
      const d = new Date(Math.round((x - 25569) * 86400000));
      return d.toISOString().slice(0, 10);
    }
    const s = String(x || "").trim();
    let m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s);
    if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
    m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
    return m ? `${m[1]}-${m[2]}-${m[3]}` : "";
  };
  const vistosLog = new Set<string>();
  const vistosNI = new Set<string>();
  const logaram: { data: string; recurso: string }[] = [];
  const ni: { data: string; recurso: string }[] = [];
  let linhas = 0, naoLogaram = 0, naoIniciaram = 0, semData = 0;
  for (let i = ini; i < v.length; i++) {
    const data = iso(v[i][cData]);
    const recurso = String(v[i][cTec] || "").replace(/^\s*Clique para ver a[cç][oõ]es\s*/i, "").trim();
    if (!data && !recurso) continue;
    linhas++;
    if (!data || !recurso) { semData++; continue; }
    const ico = cIco >= 0 ? sa(v[i][cIco]) : "";
    const k = data + "|" + sa(recurso);
    if (NAO_INICIOU.test(ico)) {
      naoIniciaram++;
      if (!vistosNI.has(k)) { vistosNI.add(k); ni.push({ data, recurso }); }
    } else if (cIco < 0 || LOGOU.test(ico)) {
      if (!vistosLog.has(k)) { vistosLog.add(k); logaram.push({ data, recurso }); }
    } else {
      naoLogaram++;
    }
  }
  // quem logou no dia não entra como "rota não iniciada"
  const naoIniciou = ni.filter(p => !vistosLog.has(p.data + "|" + sa(p.recurso)));
  return { logaram, naoIniciou, debug: { aba: nomeAba, linhas, logaram: logaram.length, naoLogaram, naoIniciaram, semData, aviso } };
}

// ===================================================================
// AVALIAÇÃO DE EXPERIÊNCIA (30/60/90 dias) → dados.json "avaliacoesExperiencia"
// Lê a aba AVALIACAO_EXPERIENCIA (gravada pelo script "Avaliação Experiência"), coluna Respostas_JSON.
// ===================================================================
function lerAvaliacoesExperiencia(workbook: ExcelScript.Workbook, aba: string = "AVALIACAO_EXPERIENCIA"): string[] {
  const ws = workbook.getWorksheet(aba);
  if (!ws) return [];
  const usado = ws.getUsedRange();
  if (!usado) return [];
  const v = usado.getValues();
  if (v.length < 2) return [];
  const cab = v[0].map(x => String(x).trim());
  const iJson = cab.indexOf("Respostas_JSON");
  const iAt = cab.indexOf("Atualizado_Em");
  if (iJson < 0) return [];
  const out: string[] = [];
  for (let r = 1; r < v.length; r++) {
    const j = String(v[r][iJson] || "").trim();
    if (!j) continue;
    // devolve o JSON como texto + quando foi gravado; o painel faz o JSON.parse
    out.push(JSON.stringify({ r: j, em: iAt >= 0 ? String(v[r][iAt] || "") : "" }));
  }
  return out;
}