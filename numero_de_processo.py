import urllib3
import requests

# Desativa o aviso de requisição HTTPS não verificada
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def gerar_num_processo_grp():
    xml = f"""<?xml version="1.0" ?>
    <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:int="http://interfaces.webservices.institucional.business.grp.prodemge.gov.br/">
    <soapenv:Header/>
    <soapenv:Body>
        <int:gerarNumeroProcessoGrp>
        <anoExercicio>2026</anoExercicio>
        <codigoOrgao>20070</codigoOrgao>
        <codigoProcessoDeNegocio>18</codigoProcessoDeNegocio>
        <codigoUnidadeAdministrativa>2007080001</codigoUnidadeAdministrativa>
        <codigoUnidadeOrcamentaria>20071</codigoUnidadeOrcamentaria>
        <codigoSistemaSolicitante>51</codigoSistemaSolicitante>
        </int:gerarNumeroProcessoGrp>
    </soapenv:Body>
    </soapenv:Envelope>"""

    return requests.post(
        "https://www.grp.servicos.mg.gov.br/serviceBus/institucional/GerarNumeroProcessoGrpWS",
        data=xml.encode("utf-8"),
        headers={"Content-Type": "text/xml; charset=utf-8"},
        timeout=90, 
        allow_redirects=False,
        # proxies=None,
        verify=False
    )


response = gerar_num_processo_grp()

start = response.text.find("<return>")
end = response.text.find("</return>")

print(f"Status Code: {response.status_code}")

if start == -1 or end == -1:
    print("ERROR")
else:
    sub = response.text[start+8:end]
    print(f"Response: {sub}")

print("\n"*2)
