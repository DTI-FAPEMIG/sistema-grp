from openpyxl import load_workbook
from datetime import datetime
import urllib3
import requests
import config 
import time

inicio = time.time()

# Desativa o aviso de requisição HTTPS não verificada
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

wb = load_workbook(config.PATH_FILE_XL)
sheet = wb["Bolsista"]

init = 52  # Linha inicial
end = sheet.max_row + 1 # Ultima linha + 1 por causa do range

for i in range(init, end):
  inicio_linha = time.time()

  now = datetime.now()
  day = now.day
  month = now.month
  year = now.year
  hour = now.hour
  minute = now.minute
  sec = now.second
  ms = f"{now.microsecond}"[3::]

  num_req = f"{day:02d}{month:02d}{now:%y}{hour:02d}{minute:02d}{sec:02d}{ms}"
  ident_proc_exec = sheet[f"AA{i}"].value
  cod_sis_origem = sheet[f"AB{i}"].value
  cod_uni_orcament = sheet[f"AD{i}"].value
  cod_org_exec = sheet[f"AE{i}"].value
  cod_uni_adm = sheet[f"A{i}"].value
  cpf_ordenador_desp = sheet[f"S{i}"].value
  justificativa = sheet[f"U{i}"].value
  tip_emp_comum = sheet[f"O{i}"].value
  ind_rec_pas = sheet[f"AF{i}"].value
  cpf_operador = sheet[f"AG{i}"].value
  num_proc_geral = sheet[f"AH{i}"].value
  cpf_cnpj_credor = sheet[f"Q{i}"].value
  classific_credor = sheet[f"P{i}"].value
  cod_fun = sheet[f"C{i}"].value
  cod_subfuncao = sheet[f"D{i}"].value
  cod_prog = sheet[f"E{i}"].value
  cod_proj_ativ = sheet[f"F{i}"].value
  cod_subproj_subativ = sheet[f"AI{i}"].value
  cod_categ_eco = sheet[f"G{i}"].value  # 3
  cod_grupo_desp = sheet[f"H{i}"].value  # 3
  cod_modal_desp = sheet[f"I{i}"].value  # 90
  cod_elemen_desp = sheet[f"J{i}"].value
  cod_item = sheet[f"AJ{i}"].value
  cod_iag = sheet[f"K{i}"].value  # 00
  cod_grupo_fonte = sheet[f"L{i}"].value  # 11
  cod_fonte_recurso = sheet[f"M{i}"].value  # 500010
  cod_ipu = sheet[f"N{i}"].value  # 01
  valor_emp = sheet[f"R{i}"].value
  cod_cc = sheet[f"T{i}"].value

  xml = f"""<?xml version="1.0" encoding="utf-8"?>
  <soap-env:Envelope xmlns:grp="http://interfaces.webservices.empenho.business.grp.prodemge.gov.br/"
                    xmlns:soap-env="http://schemas.xmlsoap.org/soap/envelope/">
    <soap-env:Body>
      <grp:gerarSolicitacaoEmpenho>
        <solicitacaoEmpenho>
          <numeroRequisicao>{num_req}</numeroRequisicao>
          <identificadorProcessoExecucao>{ident_proc_exec}</identificadorProcessoExecucao>
          <codigoSistemaOrigem>{cod_sis_origem}</codigoSistemaOrigem>
          <anoExercicioFinanceiro>{year}</anoExercicioFinanceiro>
          <codigoUnidadeOrcamentaria>{cod_uni_orcament}</codigoUnidadeOrcamentaria>
          <codigoOrgaoExecutor>{cod_org_exec}</codigoOrgaoExecutor>
          <codigoUnidadeAdministrativa>{cod_uni_adm}</codigoUnidadeAdministrativa>
          <cpfOrdenadorDespesa>{cpf_ordenador_desp}</cpfOrdenadorDespesa>
          <justificativa>{justificativa}</justificativa>
          <tipoEmpenhoComum>{tip_emp_comum}</tipoEmpenhoComum>
          <olpPendenteCreditoNaoRegularizar>
            <indicadorReconhecimentoPassivo>{ind_rec_pas}</indicadorReconhecimentoPassivo>
            <cpfOperador>{cpf_operador}</cpfOperador>
            <numeroProcessoGeral>{num_proc_geral}</numeroProcessoGeral>
            <cpfCnpjCredor>{cpf_cnpj_credor}</cpfCnpjCredor>
            <classificacaoCredor>{classific_credor}</classificacaoCredor>
            <codigoFuncao>{cod_fun}</codigoFuncao>
            <codigoSubfuncao>{cod_subfuncao}</codigoSubfuncao>
            <codigoPrograma>{cod_prog}</codigoPrograma>
            <codigoProjetoAtividade>{cod_proj_ativ}</codigoProjetoAtividade>
            <codigoSubprojetoSubatividade>{cod_subproj_subativ}</codigoSubprojetoSubatividade>
            <codigoCategoriaEconomica>{cod_categ_eco}</codigoCategoriaEconomica>
            <codigoGrupoDespesa>{cod_grupo_desp}</codigoGrupoDespesa>
            <codigoModalidadeDespesa>{cod_modal_desp}</codigoModalidadeDespesa>
            <codigoElementoDespesa>{cod_elemen_desp}</codigoElementoDespesa>
            <codigoItem>{cod_item}</codigoItem>
            <codigoIag>{cod_iag}</codigoIag>
            <codigoGrupoFonte>{cod_grupo_fonte}</codigoGrupoFonte>
            <codigoFonteRecurso>{cod_fonte_recurso}</codigoFonteRecurso>
            <codigoIpu>{cod_ipu}</codigoIpu>
            <valorEmpenho>{valor_emp}</valorEmpenho>
            <codigoCc>{cod_cc}</codigoCc>
          </olpPendenteCreditoNaoRegularizar>
        </solicitacaoEmpenho>
      </grp:gerarSolicitacaoEmpenho>
    </soap-env:Body>
  </soap-env:Envelope>"""

  response = requests.post(
      "https://www.grp.servicos.mg.gov.br/serviceBus/empenho/SolicitacaoEmpenhoWS",
      data=xml.encode("utf-8"),
      headers={"Content-Type": "text/xml; charset=utf-8"},
      timeout=90,
      allow_redirects=False,
      #proxies=None,
      verify=False
  )  # sem www o servidor responde 301: não siga sozinho

  response_text = response.text

  if response_text.find("<numeroRequisicao>") != -1:
    start_num_req = response_text.find("<numeroRequisicao>") + len("<numeroRequisicao>")
    end_num_req = response_text.find("</numeroRequisicao>")
    numero_requisicao = response_text[start_num_req : end_num_req]
    sheet[f"Z{i}"] = numero_requisicao

  if response_text.find("<numeroDocumento>") != -1:
    start_num_doc = response_text.find("<numeroDocumento>") + len("<numeroDocumento>")
    end_num_doc = response_text.find("</numeroDocumento>")
    empenho = response_text[start_num_doc : end_num_doc]
    sheet[f"W{i}"] = empenho

  # Salvando as alterações no arquivo
  wb.save(config.PATH_FILE_XL)

  fim_linha = time.time()
  total_linha = fim_linha - inicio_linha
  print(f"Linha {i} | tempo gasto: {total_linha}\n")

  # print("\n\n#############################################################\n\n")
  # print(response.status_code)
  # print("\n\n#############################################################\n\n")
  # print(response.text)
  # print("\n\n#############################################################\n\n")
  # print(response.request.body)
  # print("\n\n#############################################################\n\n")

fim = time.time()
tempo_total = fim - inicio
print(f"########\nTempo total gasto: {tempo_total}\n############")