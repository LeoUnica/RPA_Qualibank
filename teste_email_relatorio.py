"""
RPA - QualiBank: envio de e-mail de TESTE do relatorio de execucao.

Script auxiliar, fora do fluxo do RPA (nao mexe em login.py nem em
propostas.py). Gera dados 100% ficticios - nenhuma proposta real e lida ou
aprovada - so para validar o layout do e-mail e a integracao com o Outlook
antes de colocar o RPA em producao/agendamento.

Rodar com: python teste_email_relatorio.py
"""

import os
from datetime import datetime, timedelta

from openpyxl import Workbook

from email_outlook_web import enviar_email_outlook_web
from relatorio_execucao import (
    DESTINATARIO_PADRAO,
    Metricas,
    coletar_anexos,
    gerar_relatorio_html,
    registrar_metricas,
    salvar_estatisticas,
    setup_logging,
)

PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))

NOTA_TOPO = (
    "🧪 ESTE É UM E-MAIL DE TESTE — nenhuma proposta real foi lida, processada ou aprovada. "
    "Todos os contratos, clientes, lojas, valores e status abaixo são dados fictícios de exemplo, "
    "usados apenas para validar o layout deste relatório e o envio automático pelo Outlook. "
    "Em produção, este e-mail é gerado e enviado automaticamente ao final de CADA execução do RPA "
    "— a ideia é agendar o robô para rodar a cada 10 minutos, então este mesmo relatório (com dados "
    "reais e o log daquela execução) chegará nesta caixa de entrada aproximadamente a cada 10 minutos."
)


def _resultados_ficticios():
    return [
        {
            "contrato": "TESTE-0001",
            "nome": "TESTE - Cliente Fictício Um",
            "loja": "TESTE - Loja Centro",
            "liquido": 8500.00,
            "data_proposta": "10/09/2026 09:12",
            "data_aprovacao_promotora": "10/09/2026 10:45",
            "data_aprovacao_supervisor": "10/09/2026 11:03",
            "aprovado": True,
        },
        {
            "contrato": "TESTE-0002",
            "nome": "TESTE - Cliente Fictício Dois",
            "loja": "TESTE - Loja Sul",
            "liquido": 15750.30,
            "data_proposta": "10/09/2026 09:40",
            "data_aprovacao_promotora": "10/09/2026 11:20",
            "data_aprovacao_supervisor": "",
            "aprovado": False,
        },
        {
            "contrato": "TESTE-0003",
            "nome": "TESTE - Cliente Fictício Três",
            "loja": "TESTE - Loja Norte",
            "liquido": 3200.00,
            "data_proposta": "10/09/2026 10:05",
            "data_aprovacao_promotora": "10/09/2026 11:50",
            "data_aprovacao_supervisor": "10/09/2026 12:01",
            "aprovado": True,
        },
        {
            "contrato": "?",
            "nome": "TESTE - (erro simulado de leitura)",
            "loja": "TESTE - Loja Leste",
            "liquido": None,
            "data_proposta": "",
            "data_aprovacao_promotora": "",
            "data_aprovacao_supervisor": "",
            "aprovado": False,
        },
    ]


def _planilha_ficticia(caminho: str) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "TESTE"
    ws.append(["Contrato", "Cliente", "Loja", "Valor Líquido", "Status"])
    for item in _resultados_ficticios():
        ws.append(
            [
                item["contrato"],
                item["nome"],
                item["loja"],
                item["liquido"] if item["liquido"] is not None else "-",
                "Aprovado" if item["aprovado"] else ("Erro" if item["contrato"] == "?" else "Ignorado"),
            ]
        )
    wb.save(caminho)
    return caminho


def main():
    logger, caminho_log, timestamp = setup_logging(PASTA_PROJETO)
    logger.info("=== INÍCIO DO ENVIO DE TESTE (dados fictícios, nenhuma proposta real) ===")

    resultados = _resultados_ficticios()

    metricas = Metricas(
        inicio=datetime.now() - timedelta(minutes=10),
        dry_run=True,
    )
    metricas.total_encontradas = len(resultados)
    metricas.tentativas_reprocessamento = 1
    metricas.fim = datetime.now()

    stats = metricas.calcular(resultados)
    registrar_metricas(logger, stats)
    salvar_estatisticas(stats, PASTA_PROJETO, timestamp)

    caminho_planilha_teste = os.path.join(PASTA_PROJETO, "logs", f"TESTE_relatorio_{timestamp}.xlsx")
    _planilha_ficticia(caminho_planilha_teste)

    anexos = coletar_anexos(caminho_log, caminho_planilha_teste, PASTA_PROJETO)
    corpo_html = gerar_relatorio_html(stats, resultados, nota_topo=NOTA_TOPO)
    assunto = f"[Teste] [RPA Qualibank] Execução Concluída - {stats['fim_dt'].strftime('%d/%m/%Y %H:%M')}"

    print("\nAbrindo o navegador no Outlook Web. O login com a conta de automação "
          "é automático; se pedir verificação adicional (MFA), conclua na janela "
          "que vai abrir - você tem até 15 minutos.")
    enviar_email_outlook_web(
        corpo_html, anexos, assunto=assunto, destinatario=DESTINATARIO_PADRAO,
        headless=False, logger=logger,
    )
    logger.info("E-mail de TESTE enviado com sucesso via Outlook Web.")
    print(f"\nE-mail de teste enviado. Anexos: {anexos}")


if __name__ == "__main__":
    main()
