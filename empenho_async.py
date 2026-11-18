import asyncio
import time
from datetime import datetime
import aiohttp
from openpyxl import load_workbook
import config

# TODO: erros

URL = "https://www.grp.servicos.mg.gov.br/serviceBus/empenho/SolicitacaoEmpenhoWS"
HEADERS = {"Content-Type": "text/xml; charset=utf-8"}

# Limita a quantidade de requisições simultâneas para não derrubar o servidor ou tomar bloqueio
LIMITE_CONCORRENCIA = 2


def extrair_tag(xml_text: str, tag: str) -> str | None:
    abertura = f"<{tag}>"
    fechamento = f"</{tag}>"
    start = xml_text.find(abertura)
    if start != -1:
        start += len(abertura)
        end = xml_text.find(fechamento, start)
        if end != -1:
            return xml_text[start:end]
    return None


def montar_xml(sheet, row: int, num_req: str, year: int) -> str:
    return f"""<?xml version="1.0" encoding="utf-8"?>
    <soap-env:Envelope xmlns:grp="http://interfaces.webservices.empenho.business.grp.prodemge.gov.br/"
                      xmlns:soap-env="http://schemas.xmlsoap.org/soap/envelope/">
      <soap-env:Body>
        <grp:gerarSolicitacaoEmpenho>
          <solicitacaoEmpenho>
            <numeroRequisicao>{num_req}</numeroRequisicao>
            <identificadorProcessoExecucao>{sheet[f"AA{row}"].value}</identificadorProcessoExecucao>
            <codigoSistemaOrigem>{sheet[f"AB{row}"].value}</codigoSistemaOrigem>
            <anoExercicioFinanceiro>{year}</anoExercicioFinanceiro>
            <codigoUnidadeOrcamentaria>{sheet[f"AD{row}"].value}</codigoUnidadeOrcamentaria>
            <codigoOrgaoExecutor>{sheet[f"AE{row}"].value}</codigoOrgaoExecutor>
            <codigoUnidadeAdministrativa>{sheet[f"A{row}"].value}</codigoUnidadeAdministrativa>
            <cpfOrdenadorDespesa>{sheet[f"S{row}"].value}</cpfOrdenadorDespesa>
            <justificativa>{sheet[f"U{row}"].value}</justificativa>
            <tipoEmpenhoComum>{sheet[f"O{row}"].value}</tipoEmpenhoComum>
            <olpPendenteCreditoNaoRegularizar>
              <indicadorReconhecimentoPassivo>{sheet[f"AF{row}"].value}</indicadorReconhecimentoPassivo>
              <cpfOperador>{sheet[f"AG{row}"].value}</cpfOperador>
              <numeroProcessoGeral>{sheet[f"AH{row}"].value}</numeroProcessoGeral>
              <cpfCnpjCredor>{sheet[f"Q{row}"].value}</cpfCnpjCredor>
              <classificacaoCredor>{sheet[f"P{row}"].value}</classificacaoCredor>
              <codigoFuncao>{sheet[f"C{row}"].value}</codigoFuncao>
              <codigoSubfuncao>{sheet[f"D{row}"].value}</codigoSubfuncao>
              <codigoPrograma>{sheet[f"E{row}"].value}</codigoPrograma>
              <codigoProjetoAtividade>{sheet[f"F{row}"].value}</codigoProjetoAtividade>
              <codigoSubprojetoSubatividade>{sheet[f"AI{row}"].value}</codigoSubprojetoSubatividade>
              <codigoCategoriaEconomica>{sheet[f"G{row}"].value}</codigoCategoriaEconomica>
              <codigoGrupoDespesa>{sheet[f"H{row}"].value}</codigoGrupoDespesa>
              <codigoModalidadeDespesa>{sheet[f"I{row}"].value}</codigoModalidadeDespesa>
              <codigoElementoDespesa>{sheet[f"J{row}"].value}</codigoElementoDespesa>
              <codigoItem>{sheet[f"AJ{row}"].value}</codigoItem>
              <codigoIag>{sheet[f"K{row}"].value}</codigoIag>
              <codigoGrupoFonte>{sheet[f"L{row}"].value}</codigoGrupoFonte>
              <codigoFonteRecurso>{sheet[f"M{row}"].value}</codigoFonteRecurso>
              <codigoIpu>{sheet[f"N{row}"].value}</codigoIpu>
              <valorEmpenho>{sheet[f"R{row}"].value}</valorEmpenho>
              <codigoCc>{sheet[f"T{row}"].value}</codigoCc>
            </olpPendenteCreditoNaoRegularizar>
          </solicitacaoEmpenho>
        </grp:gerarSolicitacaoEmpenho>
      </soap-env:Body>
    </soap-env:Envelope>"""


async def processar_linha(session: aiohttp.ClientSession, semaphore: asyncio.Semaphore, row: int, xml: str):
    async with semaphore:
        inicio_linha = time.time()
        try:
            async with session.post(
                URL,
                data=xml.encode("utf-8"),
                headers=HEADERS,
                allow_redirects=False,
                ssl=False  # Equivalente ao verify=False
            ) as response:
                response_text = await response.text()

                # ----------------------------------------
                print(f"\n\n################ Linha {row} #############################\n")
                print(f"Status Code: {response.status}")
                print("\n---------------- Corpo da Resposta (response.text) -------------\n")
                print(response_text)
                print("\n---------------- XML Enviado (request.body) ------------------\n")
                print(xml)
                print("\n#############################################################\n")
                # ----------------------------------------

                num_req_retornado = extrair_tag(response_text, "numeroRequisicao")
                empenho_retornado = extrair_tag(response_text, "numeroDocumento")

                tempo = time.time() - inicio_linha
                print(f"Linha {row} concluída | tempo gasto: {tempo:.2f}s")
                return row, num_req_retornado, empenho_retornado

        except Exception as e:
            print(f"Erro na linha {row}: {e}")
            return row, None, None


async def main():
    inicio = time.time()

    wb = load_workbook(config.PATH_FILE_XL)
    sheet = wb["Bolsista"]

    init = 52
    end = sheet.max_row + 1

    semaphore = asyncio.Semaphore(LIMITE_CONCORRENCIA)
    timeout = aiohttp.ClientTimeout(total=90)

    # Cria a sessão HTTP reutilizável (Connection Pooling)
    async with aiohttp.ClientSession(timeout=timeout) as session:
        tarefas = []

        for idx, i in enumerate(range(init, end)):
            now = datetime.now()
            # Adicionado o índice (idx) para evitar colisão de numeroRequisicao gerado no mesmo milissegundo
            ms = f"{now.microsecond:06d}"[3:]
            num_req = f"{now:%d%m%y%H%M%S}{ms}{idx:03d}"

            xml = montar_xml(sheet, i, num_req, now.year)
            tarefas.append(processar_linha(session, semaphore, i, xml))

        # Dispara todas as requisições de forma concorrente
        resultados = await asyncio.gather(*tarefas)

    # Atualiza a planilha na memória de forma segura
    for row, num_req, empenho in resultados:
        if num_req:
            sheet[f"Z{row}"] = num_req
        if empenho:
            sheet[f"W{row}"] = empenho

    # Salva no disco apenas UMA vez no final
    wb.save(config.PATH_FILE_XL)

    tempo_total = time.time() - inicio
    print(f"########\nTempo total gasto: {tempo_total:.2f}s\n############")


if __name__ == "__main__":
    asyncio.run(main())
