"""
RPA - QualiBank: observabilidade da execucao.

Modulo de suporte responsavel por logging estruturado, consolidacao de
metricas executivas, geracao do relatorio HTML e envio do e-mail de
encerramento via Microsoft Outlook (COM). Nao contem nenhuma regra de
negocio de aprovacao/liberacao de propostas - isso permanece exclusivamente
em propostas.py.
"""

import json
import logging
import os
import socket
import subprocess
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional

DESTINATARIO_PADRAO = "unica.tech@unicapromotora.com.br"
MAX_LINHAS_TABELA_EMAIL = 200
VERSAO_RELATORIO = "1.0.0"


# ======================================================================
# 1. LOGGING
# ======================================================================


def setup_logging(pasta_projeto: str) -> tuple[logging.Logger, str, str]:
    """Configura o logger raiz da automacao ("rpa_qualibank").

    Cria (se necessario) a pasta logs/ e um arquivo
    logs/automacao_YYYYMMDD_HHMMSS.log. Modulos individuais (login.py,
    propostas.py) devem usar logging.getLogger("rpa_qualibank.<modulo>"),
    que propaga para os handlers configurados aqui.

    Retorna (logger, caminho_do_log, timestamp_da_execucao).
    """
    pasta_logs = os.path.join(pasta_projeto, "logs")
    os.makedirs(pasta_logs, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho_log = os.path.join(pasta_logs, f"automacao_{timestamp}.log")

    logger = logging.getLogger("rpa_qualibank")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()
    logger.propagate = False

    formato = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.FileHandler(caminho_log, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formato)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formato)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger, caminho_log, timestamp


# ======================================================================
# 1b. LOCK DE EXECUCAO (evita duas execucoes simultaneas / dois logins)
# ======================================================================

NOME_ARQUIVO_LOCK = "rpa.lock"


def _pid_esta_rodando(pid: int) -> bool:
    """Verifica, via tasklist (Windows), se um PID ainda esta em execucao."""
    try:
        resultado = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        return str(pid) in resultado.stdout
    except Exception:
        # Se nao for possivel checar, assume que nao esta rodando - melhor
        # arriscar uma execucao duplicada rara do que travar o robo para sempre
        # por causa de um lock que nunca mais consegue ser validado.
        return False


def adquirir_lock_execucao(pasta_projeto: str) -> bool:
    """Garante que so uma execucao do RPA rode por vez.

    O agendador dispara uma nova execucao a cada 10 minutos, mas se a
    anterior ainda estiver rodando (ex.: travada aguardando MFA ou uma
    pagina lenta), a nova execucao NAO deve logar de nova na mesma conta ao
    mesmo tempo. Retorna True se conseguiu o lock (pode prosseguir com login
    normalmente) ou False se ja ha uma execucao em andamento (deve encerrar
    sem tentar logar)."""
    caminho = os.path.join(pasta_projeto, NOME_ARQUIVO_LOCK)

    if os.path.exists(caminho):
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                pid_anterior = int(f.read().strip())
        except (ValueError, OSError):
            pid_anterior = None

        if pid_anterior is not None and _pid_esta_rodando(pid_anterior):
            return False
        # Lock "orfao" de uma execucao anterior que travou/crashou sem
        # limpar o proprio lock - remove e segue normalmente.

    with open(caminho, "w", encoding="utf-8") as f:
        f.write(str(os.getpid()))
    return True


def liberar_lock_execucao(pasta_projeto: str) -> None:
    """Remove o lock de execucao. Deve ser chamado sempre no final (bloco
    finally), sucesso ou falha, para nao deixar a proxima execucao presa."""
    caminho = os.path.join(pasta_projeto, NOME_ARQUIVO_LOCK)
    try:
        os.remove(caminho)
    except OSError:
        pass


# ======================================================================
# 2. METRICAS EXECUTIVAS
# ======================================================================


@dataclass
class Metricas:
    """Acumula os poucos dados que nao podem ser derivados da lista de
    resultados (inicio, total encontrado, tentativas de reprocessamento,
    erro critico). Tudo o mais e calculado a partir de `resultados` em
    calcular()."""

    inicio: datetime = field(default_factory=datetime.now)
    fim: Optional[datetime] = None
    total_encontradas: int = 0
    tentativas_reprocessamento: int = 0
    erro_critico: Optional[str] = None
    dry_run: bool = False

    def calcular(self, resultados: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Consolida as metricas executivas a partir da lista de resultados
        produzida por processar_propostas()."""
        fim = self.fim or datetime.now()

        total_processadas = len(resultados)
        total_erro = sum(1 for item in resultados if item.get("contrato") == "?")
        total_aprovadas = sum(1 for item in resultados if item.get("aprovado"))
        total_nao_aprovadas = total_processadas - total_aprovadas - total_erro

        valores_validos = [
            item["liquido"]
            for item in resultados
            if isinstance(item.get("liquido"), (int, float))
        ]
        valores_aprovados = [
            item["liquido"]
            for item in resultados
            if item.get("aprovado") and isinstance(item.get("liquido"), (int, float))
        ]

        valor_total_aprovado = sum(valores_aprovados)
        valor_medio = (sum(valores_validos) / len(valores_validos)) if valores_validos else 0.0
        maior_valor = max(valores_validos) if valores_validos else 0.0
        menor_valor = min(valores_validos) if valores_validos else 0.0

        processadas_sem_erro = total_processadas - total_erro
        taxa_sucesso = (processadas_sem_erro / total_processadas * 100) if total_processadas else 0.0

        return {
            "inicio": self.inicio.strftime("%d/%m/%Y %H:%M:%S"),
            "fim": fim.strftime("%d/%m/%Y %H:%M:%S"),
            "fim_dt": fim,
            "tempo_execucao": str(fim - self.inicio).split(".")[0],
            "total_encontradas": self.total_encontradas,
            "total_processadas": total_processadas,
            "total_aprovadas": total_aprovadas,
            "total_nao_aprovadas": total_nao_aprovadas,
            "total_erro": total_erro,
            "valor_total_aprovado": valor_total_aprovado,
            "valor_medio": valor_medio,
            "maior_valor": maior_valor,
            "menor_valor": menor_valor,
            "taxa_sucesso": round(taxa_sucesso, 2),
            "tentativas_reprocessamento": self.tentativas_reprocessamento,
            "erro_critico": self.erro_critico,
            "dry_run": self.dry_run,
        }


def registrar_metricas(logger: logging.Logger, stats: Dict[str, Any]) -> None:
    """Registra no log, em nivel INFO, o resumo executivo da execucao."""
    logger.info("===== RESUMO DA EXECUCAO =====")
    logger.info(f"Inicio: {stats['inicio']} | Fim: {stats['fim']} | Duracao: {stats['tempo_execucao']}")
    logger.info(f"Propostas encontradas: {stats['total_encontradas']}")
    logger.info(f"Propostas processadas: {stats['total_processadas']}")
    logger.info(f"Aprovadas: {stats['total_aprovadas']}")
    logger.info(f"Nao aprovadas: {stats['total_nao_aprovadas']}")
    logger.info(f"Com erro: {stats['total_erro']}")
    logger.info(f"Valor total aprovado: {_fmt_moeda(stats['valor_total_aprovado'])}")
    logger.info(f"Valor medio das propostas: {_fmt_moeda(stats['valor_medio'])}")
    logger.info(f"Maior valor encontrado: {_fmt_moeda(stats['maior_valor'])}")
    logger.info(f"Menor valor encontrado: {_fmt_moeda(stats['menor_valor'])}")
    logger.info(f"Taxa de sucesso: {stats['taxa_sucesso']}%")
    logger.info(f"Tentativas de reprocessamento: {stats['tentativas_reprocessamento']}")
    if stats["erro_critico"]:
        logger.critical(f"Erro critico registrado na execucao: {stats['erro_critico']}")
    logger.info("===============================")


def salvar_estatisticas(stats: Dict[str, Any], pasta_projeto: str, timestamp: str) -> str:
    """Persiste as metricas executivas em logs/estatisticas_<timestamp>.json,
    para fins de auditoria/historico. Retorna o caminho do arquivo salvo."""
    pasta_logs = os.path.join(pasta_projeto, "logs")
    os.makedirs(pasta_logs, exist_ok=True)
    caminho = os.path.join(pasta_logs, f"estatisticas_{timestamp}.json")

    stats_serializaveis = {k: v for k, v in stats.items() if k != "fim_dt"}
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(stats_serializaveis, f, ensure_ascii=False, indent=2)

    return caminho


# ======================================================================
# 3. RELATORIO HTML
# ======================================================================


def _fmt_moeda(valor: Optional[float]) -> str:
    """Formata um numero como moeda brasileira, ex.: R$ 1.234,56."""
    if valor is None:
        return "-"
    texto = f"{valor:,.2f}"
    texto = texto.replace(",", "_").replace(".", ",").replace("_", ".")
    return f"R$ {texto}"


# O editor de composicao do Outlook Web nao preserva estilo CSS custom
# (background-color, border, border-radius, cor de texto via style) ao
# enviar - ele serializa a partir do seu proprio modelo interno, que so
# reconhece formatacao aplicada via comandos reais do editor (negrito,
# italico, listas). Confirmado por teste: <b>/<i>/<ul><li> sobrevivem,
# style="color:..."/"background-color:..." nao. Por isso o relatorio usa
# so essas tags "seguras" + emoji (que sao caracteres, nao CSS, e sempre
# aparecem coloridos) para comunicar status.
_STATUS_EMOJI = {
    "aprovado": "✅",
    "ignorado": "⚠️",
    "erro": "❌",
}


def _status_proposta(item: Dict[str, Any]) -> tuple[str, str]:
    """Retorna (texto_status, emoji) para uma linha da tabela."""
    if item.get("contrato") == "?":
        return "Erro", _STATUS_EMOJI["erro"]
    if item.get("aprovado"):
        return "Aprovado", _STATUS_EMOJI["aprovado"]
    return "Ignorado", _STATUS_EMOJI["ignorado"]


def status_execucao(stats: Dict[str, Any]) -> tuple[str, str, str]:
    """Classifica o resultado geral da execucao a partir das metricas.

    Retorna (nivel, texto_curto, emoji). Usado tanto no assunto do e-mail
    quanto na linha de status do corpo, para os dois ficarem sempre
    coerentes entre si."""
    if stats.get("erro_critico"):
        return "erro", "FALHOU", "❌"
    if stats.get("total_erro", 0) > 0:
        return "pendencias", "CONCLUÍDA COM PENDÊNCIAS", "⚠️"
    return "sucesso", "CONCLUÍDA COM SUCESSO", "✅"


def _passos_execucao(stats: Dict[str, Any]) -> List[tuple[str, bool]]:
    """Deriva quais etapas do fluxo foram concluidas a partir das metricas
    disponiveis (nao existe rastreamento de etapa a etapa no robo hoje -
    isso e so a melhor inferencia a partir do que ja temos)."""
    erro_critico = bool(stats.get("erro_critico"))
    tem_dados = stats.get("total_encontradas", 0) > 0
    processou = stats.get("total_processadas", 0) > 0
    return [
        ("Leitura da base", tem_dados or not erro_critico),
        ("Processamento", processou),
        ("Validação", processou),
        ("Aprovação", processou and not erro_critico),
        ("Envio do relatório", True),
    ]


def gerar_relatorio_html(
    stats: Dict[str, Any],
    resultados: List[Dict[str, Any]],
    nota_topo: Optional[str] = None,
) -> str:
    """Monta o corpo HTML do e-mail de encerramento da execucao.

    `nota_topo`, se informado, e exibido como um aviso azul de destaque logo
    no topo do corpo do e-mail (ex.: avisar que o envio e um teste com dados
    ficticios). Nao interfere no calculo das metricas nem na tabela."""
    linhas_html = []
    for item in resultados[:MAX_LINHAS_TABELA_EMAIL]:
        status_texto, emoji_status = _status_proposta(item)
        linhas_html.append(
            "<tr>"
            f'<td><b>{item.get("contrato") or "-"}</b></td>'
            f'<td>{item.get("nome") or "-"}</td>'
            f'<td>{item.get("loja") or "-"}</td>'
            f'<td align="right">{_fmt_moeda(item.get("liquido"))}</td>'
            f'<td>{item.get("data_proposta") or "-"}</td>'
            f'<td>{item.get("data_aprovacao_promotora") or "-"}</td>'
            f'<td>{item.get("data_aprovacao_supervisor") or "-"}</td>'
            f'<td align="center">{emoji_status} {status_texto}</td>'
            "</tr>"
        )

    nota_truncamento = ""
    if len(resultados) > MAX_LINHAS_TABELA_EMAIL:
        restante = len(resultados) - MAX_LINHAS_TABELA_EMAIL
        nota_truncamento = (
            f"<p>Exibindo as primeiras {MAX_LINHAS_TABELA_EMAIL} propostas nesta tabela "
            f"({restante} adicionais no relatório Excel em anexo).</p>"
        )

    aviso_nota_topo = f"<p>{nota_topo}</p>" if nota_topo else ""

    aviso_erro_critico = ""
    if stats.get("erro_critico"):
        aviso_erro_critico = f"<p>❌ <b>A execução foi interrompida por um erro crítico:</b> {stats['erro_critico']}</p>"

    modo_execucao = "SIMULAÇÃO (DRY-RUN)" if stats.get("dry_run") else "PRODUÇÃO (aprovação real)"
    ambiente = "Simulação (Dry-Run)" if stats.get("dry_run") else "Produção"
    maquina = socket.gethostname()
    _, status_texto_geral, emoji_status_geral = status_execucao(stats)

    resumo_itens = "".join(
        [
            f"<li>✅ <b>Processadas:</b> {stats['total_processadas']} (de {stats['total_encontradas']} encontradas)</li>",
            f"<li>🟢 <b>Aprovadas:</b> {stats['total_aprovadas']}</li>",
            f"<li>🔴 <b>Não Aprovadas:</b> {stats['total_nao_aprovadas']}</li>",
            f"<li>⚠️ <b>Erros:</b> {stats['total_erro']}</li>",
            f"<li>💰 <b>Valor Aprovado:</b> {_fmt_moeda(stats['valor_total_aprovado'])}</li>",
        ]
    )

    info_itens = "".join(
        [
            f"<li>🖥️ <b>Máquina:</b> {maquina}</li>",
            f"<li>⚙️ <b>Modo:</b> {modo_execucao}</li>",
            f"<li>🕒 <b>Início:</b> {stats['inicio']}</li>",
            f"<li>🏁 <b>Fim:</b> {stats['fim']}</li>",
            f"<li>⏱️ <b>Duração:</b> {stats['tempo_execucao']}</li>",
            f"<li>🔁 <b>Reprocessamentos:</b> {stats['tentativas_reprocessamento']}</li>",
        ]
    )

    passos = _passos_execucao(stats)
    timeline_itens = "".join(
        f"<li>{'✅' if concluido else '➖'} {titulo}</li>" for titulo, concluido in passos
    )

    return f"""
<p><b>🤖 RPA Qualibank</b><br/>
<i>Relatório automático de execução</i></p>

{aviso_nota_topo}
{aviso_erro_critico}

<p><b>{emoji_status_geral} Execução {status_texto_geral.lower()}</b><br/>
{stats['total_processadas']} propostas processadas, {stats['total_aprovadas']} aprovadas, {stats['total_nao_aprovadas']} não aprovadas, {stats['total_erro']} com erro.</p>

<p><b>Resumo Executivo</b></p>
<ul>{resumo_itens}</ul>

<p><b>Informações da Execução</b></p>
<ul>{info_itens}</ul>

<p><b>Fluxo da Execução</b></p>
<ul>{timeline_itens}</ul>

<p><b>Indicadores Financeiros</b></p>
<ul>
<li>Taxa de Sucesso: <b>{stats['taxa_sucesso']}%</b></li>
<li>Valor Médio das Propostas: <b>{_fmt_moeda(stats['valor_medio'])}</b></li>
<li>Maior Valor Encontrado: <b>{_fmt_moeda(stats['maior_valor'])}</b></li>
<li>Menor Valor Encontrado: <b>{_fmt_moeda(stats['menor_valor'])}</b></li>
</ul>

<p><b>Detalhamento das Propostas</b></p>
<table border="1" cellpadding="6" cellspacing="0">
<tr>
<th align="left">Contrato</th>
<th align="left">Cliente</th>
<th align="left">Loja</th>
<th align="right">Valor Líquido</th>
<th align="left">Data da Proposta</th>
<th align="left">Data Aprovação Promotora</th>
<th align="left">Data Aprovação Supervisor</th>
<th align="center">Status</th>
</tr>
{''.join(linhas_html) if linhas_html else '<tr><td colspan="8" align="center">Nenhuma proposta processada.</td></tr>'}
</table>
{nota_truncamento}

<p><i>RPA Qualibank &mdash; Automação Inteligente de Processos</i><br/>
Este relatório foi gerado automaticamente pelo robô de processamento da Única Promotora.<br/>
Ambiente: {ambiente} &middot; Máquina: {maquina} &middot; Versão: {VERSAO_RELATORIO}<br/>
<b>Mensagem automática. Não responda este e-mail.</b></p>
"""


# ======================================================================
# 4. ANEXOS
# ======================================================================


def coletar_anexos(
    caminho_log: Optional[str],
    caminho_relatorio: Optional[str],
    pasta_projeto: str,
    desde: Optional[datetime] = None,
) -> List[str]:
    """Reune os caminhos dos arquivos a anexar no e-mail: log da execucao,
    planilha de resultado (simulacao ou aprovacoes reais) e todos os
    screenshots produzidos em screenshots/.

    Se `desde` for informado, so entram os screenshots gerados a partir
    desse instante (o inicio do ciclo atual), para o e-mail nao carregar
    evidencias de ciclos anteriores."""
    anexos: List[str] = []

    if caminho_log and os.path.isfile(caminho_log):
        anexos.append(caminho_log)

    if caminho_relatorio and os.path.isfile(caminho_relatorio):
        anexos.append(caminho_relatorio)

    pasta_screenshots = os.path.join(pasta_projeto, "screenshots")
    if os.path.isdir(pasta_screenshots):
        for nome in sorted(os.listdir(pasta_screenshots)):
            if not nome.lower().endswith(".png"):
                continue
            caminho = os.path.join(pasta_screenshots, nome)
            if desde is not None and datetime.fromtimestamp(os.path.getmtime(caminho)) < desde:
                continue
            anexos.append(caminho)

    return anexos


# ======================================================================
# 5. ENVIO PELO MICROSOFT OUTLOOK
# ======================================================================


def enviar_email_outlook(
    corpo_html: str,
    anexos: List[str],
    assunto: Optional[str] = None,
    destinatario: str = DESTINATARIO_PADRAO,
) -> None:
    """Envia o relatorio de execucao por e-mail usando o Microsoft Outlook
    instalado na maquina (integracao COM via win32com.client).

    Requer o Outlook desktop configurado e uma conta ativa na sessao do
    Windows em que o script roda.
    """
    import win32com.client as win32

    if assunto is None:
        assunto = f"[RPA Supervisor] Execução Concluída - {datetime.now().strftime('%d/%m/%Y %H:%M')}"

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = destinatario
    mail.Subject = assunto
    mail.HTMLBody = corpo_html

    for caminho in anexos:
        if os.path.isfile(caminho):
            mail.Attachments.Add(os.path.abspath(caminho))

    mail.Send()
