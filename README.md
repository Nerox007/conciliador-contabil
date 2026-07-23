Conciliador Bancário Inteligente
Como rodar
Bash
Login de demonstração: usuário demo, senha galvao3dias.
Configurar login de verdade (antes de usar com qualquer cliente)
Crie o arquivo .streamlit/secrets.toml (não versionar no Git):
Toml
Gere o hash de cada senha com:
Bash
Sem esse arquivo, o app cai automaticamente no usuário de demonstração — e mostra um aviso
amarelo na tela avisando que está em modo demo, pra ninguém esquecer de trocar.
O que foi corrigido em relação à versão original
Senha fixa no código-fonte → agora compara hash (nunca texto puro) e busca as
credenciais em st.secrets, fora do código. Suporta múltiplos usuários.
Bug do cruzamento com lançamentos duplicados: o pd.merge original, ao encontrar
duas transações com a mesma Data+Valor de cada lado, gerava produto cartesiano
(2 pares reais viravam 4 "conciliados"). Testei com dados reais: 3 pares corretos
contando 5 no código antigo. Corrigido com um contador de ocorrência por combinação
Data+Valor antes do merge.
Datas ambíguas: pd.to_datetime sem dayfirst=True lê "03/04/2026" como
4 de março (padrão americano) em vez de 3 de abril. Crítico para dados brasileiros
vindos como texto.
Valores em texto: agora trata R$ 1.234,56, 1.234,56, 1,234.56, (150,00)
como negativo, milhar sem decimais (2.500 → 2500), etc.
Detecção de colunas: se duas colunas contêm "valor" (ex.: "Valor Bruto" e
"Valor Líquido"), o app agora pede pra você escolher em vez de renomear as duas
silenciosamente e quebrar.
Novo: tolerância de dias — lançamentos com mesmo valor mas data diferindo em
até N dias (configurável na barra lateral) entram como "match aproximado" em vez
de virar divergência — resolve o caso clássico de compensação bancária.
Novo: comparação por valor absoluto (opcional) — para quando um lado registra
débito como negativo e o outro como positivo.
Relatório Excel: agora com aba de resumo executivo, cabeçalhos coloridos,
largura de coluna automática e fonte profissional — pronto pra mandar ao cliente
sem editar nada.
Diagnóstico de linhas ignoradas: quando uma linha é descartada por Data/Valor
inválidos, agora aparece um painel mostrando o número exato da linha na planilha
e o motivo — em vez de só um contador genérico.
Melhorias desta rodada (modelos, prévia, progresso)
Modelos de planilha pra baixar: expander no topo do app com dois botões
("Modelo de Extrato Bancário" e "Modelo de Sistema") que geram exemplos .xlsx
prontos, no formato certo. Os dados de exemplo foram escolhidos de propósito
para, se usados juntos, já mostrarem na prática os três tipos de match (exato,
por tolerância de dias e por valor absoluto) — uma pequena demonstração viva
do produto embutida no próprio modelo.
Prévia com estatísticas pós-upload: depois do mapeamento de colunas e antes
de clicar em "Processar", o app mostra quantas linhas foram lidas, quantas vão
ser ignoradas, o período de datas e a soma dos valores — pra você confirmar que
o mapeamento pegou as colunas certas antes de rodar a conciliação completa.
Barra de progresso real: trocei o spinner genérico por uma barra com estágios
(limpando extrato → limpando sistema → cruzando lançamentos → gerando relatório).
A etapa de cruzamento reporta progresso de verdade, calculado a partir da posição
no laço de tolerância — não é uma animação decorativa desligada do processamento.
Toda a lógica de limpeza de valores, deteção de colunas e conciliação (inclusive os
casos de duplicata e tolerância) tem testes automatizados que rodei antes de entregar.
Isso já é um SaaS?
Ainda não — e prefiro te falar isso claramente em vez de deixar parecer que está pronto:
Autenticação: o esquema atual (hash + secrets.toml) é sério o suficiente pra uso
interno ou com poucos clientes que você mesmo cadastra manualmente. Não tem cadastro
próprio, recuperação de senha, nem isolamento real de dados entre contas — é login
compartilhado, não multi-tenant.
Sem persistência: cada conciliação vive só na sessão do navegador. Não há
histórico, não há banco de dados. Pra vender de verdade, o próximo passo é um
backend (Postgres via Supabase, por exemplo) com uma conta por cliente e os dados
de cada um isolados.
Sem cobrança: nenhuma integração de pagamento (Stripe, etc.).
Streamlit tem login nativo via OAuth (st.login) que pode valer a pena migrar
para quando o número de clientes crescer — resolve cadastro/sessão sem reinventar a roda.
Pra validar a ideia com os primeiros clientes reais, o que está aqui já resolve bem.
Pra cobrar mensalidade de várias contabilidades ao mesmo tempo, o próximo investimento
é multi-tenancy de verdade.