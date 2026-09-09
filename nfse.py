"""Leitura conservadora de NFS-e nacional e ABRASF. Não valida autorização fiscal."""
from datetime import date
from decimal import Decimal, InvalidOperation
from io import BytesIO
import re
import unicodedata

from defusedxml import ElementTree as ET
from pypdf import PdfReader

MONEY_FIELDS = ('service_value', 'deductions', 'iss', 'pis', 'cofins', 'inss', 'ir', 'csll', 'net_value')
COLUMNS = {
    'client_name': 'Cliente', 'direction': 'Tipo', 'number': 'Número da nota',
    'access_key': 'Chave de acesso', 'verification_code': 'Código de verificação',
    'issued_at': 'Emissão', 'competence': 'Competência',
    'provider_name': 'Prestador', 'provider_document': 'CPF/CNPJ do prestador',
    'recipient_name': 'Tomador', 'recipient_document': 'CPF/CNPJ do tomador',
    'municipality': 'Município de emissão', 'service_code': 'Código do serviço',
    'description': 'Descrição do serviço', 'service_value': 'Valor dos serviços',
    'deductions': 'Deduções', 'iss': 'ISS', 'iss_withheld': 'ISS retido',
    'pis': 'PIS', 'cofins': 'COFINS', 'inss': 'INSS', 'ir': 'IR', 'csll': 'CSLL',
    'net_value': 'Valor líquido', 'status': 'Situação informada',
    'review_status': 'Conferência', 'source': 'Origem', 'notes': 'Observações',
}
DEFAULT_COLUMNS = ['client_name', 'direction', 'number', 'issued_at', 'competence',
                   'provider_name', 'recipient_name', 'description', 'service_value', 'iss', 'net_value']
CITIES = {'3159605': 'Santa Rita do Sapucaí / MG', '3152501': 'Pouso Alegre / MG', '3132404': 'Itajubá / MG'}


def digits(value):
    return re.sub(r'\D', '', value or '')


def document(value):
    # CNPJ alfanumérico: preservar letras; remover apenas formatação.
    return re.sub(r'[^A-Z0-9]', '', str(value or '').upper())


def money(value):
    if value is None or str(value).strip() == '':
        return None
    raw = str(value).strip().replace('R$', '').replace(' ', '')
    if ',' in raw:
        raw = raw.replace('.', '').replace(',', '.')
    try:
        result = Decimal(raw)
        if not result.is_finite() or abs(result) > Decimal('999999999999.99'):
            raise ValueError('Valor monetário fora do limite.')
        return format(result.quantize(Decimal('.01')), 'f')
    except InvalidOperation as exc:
        raise ValueError('Valor monetário inválido.') from exc


def iso_date(value):
    if not value:
        return ''
    raw = str(value).strip()
    if re.match(r'^\d{2}/\d{2}/\d{4}', raw):
        d, m, y = raw[:10].split('/')
        raw = f'{y}-{m}-{d}'
    elif re.fullmatch(r'\d{4}-\d{2}', raw):
        raw += '-01'
    return date.fromisoformat(raw[:10]).isoformat()


def node(root, *paths):
    if root is None:
        return None
    for path in paths:
        result = root.find(path)
        if result is not None:
            return result
    return None


def value(root, *paths):
    result = node(root, *paths)
    return ''.join(result.itertext()).strip() if result is not None else ''


def empty_invoice():
    item = {key: '' for key in COLUMNS if key != 'client_name'}
    item.update({key: None for key in MONEY_FIELDS})
    item.update(status='Não verificada', review_status='Pendente', direction='A conferir', notes='', raw_text='')
    return item


def parse_xml(data):
    root = ET.fromstring(data, forbid_dtd=True, forbid_entities=True, forbid_external=True)
    for element in root.iter():
        element.tag = element.tag.rsplit('}', 1)[-1]
    # Only finalized invoice containers, never DPS/RPS requests or event envelopes.
    records = [e for e in root.iter() if e.tag in ('infNFSe', 'InfNfse')]
    if not records:
        raise ValueError('XML não reconhecido como NFS-e nacional ou ABRASF. DPS/RPS e eventos isolados não são notas.')
    if len(records) > 500:
        raise ValueError('O XML excede o limite de 500 notas.')
    invoices = []
    for record in records:
        item = empty_invoice()
        if record.tag == 'infNFSe':
            dps = node(record, 'DPS/infDPS')
            if dps is None:
                raise ValueError('NFS-e nacional sem DPS incorporada; conteúdo incompleto.')
            provider = node(record, 'emit')
            if provider is None:
                provider = node(dps, 'prest')
            recipient = node(dps, 'toma')
            identifier = record.get('Id', '')
            key = identifier[3:] if identifier.startswith('NFS') else identifier
            item.update(
                number=value(record, 'nNFSe'), access_key=key,
                issued_at=iso_date(value(record, 'dhProc') or value(dps, 'dhEmi')),
                competence=iso_date(value(dps, 'dCompet')),
                provider_name=value(provider, 'xNome'), provider_document=value(provider, 'CNPJ', 'CPF'),
                recipient_name=value(recipient, 'xNome'), recipient_document=value(recipient, 'CNPJ', 'CPF'),
                municipality=value(record, 'xLocEmi') or CITIES.get(value(dps, 'cLocEmi'), value(dps, 'cLocEmi')),
                service_code=value(dps, 'serv/cServ/cTribNac'), description=value(dps, 'serv/cServ/xDescServ'),
                service_value=money(value(dps, 'valores/vServPrest/vServ')),
                deductions=money(value(dps, 'valores/vDedRed/vDR')),
                iss=money(value(record, 'valores/vISSQN')), net_value=money(value(record, 'valores/vLiq')),
                pis=money(value(dps, 'valores/trib/tribFed/piscofins/vPis')),
                cofins=money(value(dps, 'valores/trib/tribFed/piscofins/vCofins')),
                inss=money(value(dps, 'valores/trib/tribFed/vRetCP')),
                ir=money(value(dps, 'valores/trib/tribFed/vRetIRRF')),
                csll=money(value(dps, 'valores/trib/tribFed/vRetCSLL')),
                source='XML nacional',
            )
            withheld = value(dps, 'valores/trib/tribMun/tpRetISSQN')
            item['iss_withheld'] = {'1': 'Não', '2': 'Sim', '3': 'Sim'}.get(withheld, '')
        else:
            declaration = node(record, 'DeclaracaoPrestacaoServico/InfDeclaracaoPrestacaoServico')
            service = node(declaration, 'Servico') if declaration is not None else node(record, 'Servico')
            provider = node(record, 'PrestadorServico')
            recipient = node(declaration, 'TomadorServico', 'Tomador') if declaration is not None else node(record, 'TomadorServico')
            item.update(
                number=value(record, 'Numero'), verification_code=value(record, 'CodigoVerificacao'),
                issued_at=iso_date(value(record, 'DataEmissao')),
                competence=iso_date(value(declaration, 'Competencia') or value(record, 'Competencia')),
                provider_name=value(provider, 'RazaoSocial'),
                provider_document=value(provider, 'IdentificacaoPrestador/CpfCnpj/Cnpj', 'IdentificacaoPrestador/CpfCnpj/Cpf', 'IdentificacaoPrestador/Cnpj'),
                recipient_name=value(recipient, 'RazaoSocial'),
                recipient_document=value(recipient, 'IdentificacaoTomador/CpfCnpj/Cnpj', 'IdentificacaoTomador/CpfCnpj/Cpf'),
                municipality=value(record, 'OrgaoGerador/CodigoMunicipio') or value(provider, 'Endereco/CodigoMunicipio'),
                description=value(service, 'Discriminacao'), service_code=value(service, 'ItemListaServico'),
                source='XML ABRASF',
            )
            item['municipality'] = CITIES.get(item['municipality'], item['municipality'])
            tags = {'service_value': 'ValorServicos', 'deductions': 'ValorDeducoes', 'iss': 'ValorIss',
                    'pis': 'ValorPis', 'cofins': 'ValorCofins', 'inss': 'ValorInss', 'ir': 'ValorIr',
                    'csll': 'ValorCsll', 'net_value': 'ValorLiquidoNfse'}
            for field, tag in tags.items():
                item[field] = money(value(record, f'ValoresNfse/{tag}') or value(service, f'Valores/{tag}'))
            item['iss_withheld'] = {'1': 'Sim', '2': 'Não'}.get(value(service, 'IssRetido', 'Valores/IssRetido'), '')
            parent = next((e for e in root.iter('CompNfse') if record in list(e.iter())), None)
            if parent is not None and node(parent, 'NfseCancelamento', 'CancelamentoNfse') is not None:
                item['status'] = 'Cancelada'
        item['provider_document'] = document(item['provider_document'])
        item['recipient_document'] = document(item['recipient_document'])
        if not item['number'] or not item['provider_document']:
            raise ValueError('XML sem número da NFS-e ou identificação do prestador.')
        invoices.append(item)
    return invoices


def parse_pdf(data):
    reader = PdfReader(BytesIO(data))
    if reader.is_encrypted:
        raise ValueError('PDF protegido por senha. Envie uma cópia desbloqueada.')
    if len(reader.pages) > 30:
        raise ValueError('PDF com mais de 30 páginas. Importe uma nota por PDF.')
    chunks = []
    for page in reader.pages:
        contents = page.get_contents()
        if contents and len(contents.get_data()) > 8 * 1024 * 1024:
            raise ValueError('Conteúdo da página PDF excede o limite de processamento.')
        chunks.append(page.extract_text() or '')
    text = '\n'.join(chunks)[:100000]
    plain = ''.join(c for c in unicodedata.normalize('NFKD', text) if not unicodedata.combining(c))
    item = empty_invoice()
    item.update(source='PDF', raw_text=text,
                notes='PDF: confira os dados com o documento original antes de marcar como conferido.')
    if not text.strip():
        item['notes'] = 'PDF sem texto extraível. OCR não incluído: preencha os campos manualmente ou importe o XML.'
    # Only unambiguous, labelled fields. Never infer issuer/recipient from arbitrary document order.
    patterns = {
        'number': r'(?:Numero (?:da (?:NFS-e|nota)|da NFS-e)|NFS-e n[ºo.]*)\s*[:\-]?\s*(\d+)\b',
        'issued_at': r'(?:Data (?:e Hora )?(?:da )?emissao)\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})',
        'competence': r'Competencia(?: da NFS-e)?\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})',
        'service_value': r'Valor (?:total dos servicos|dos servicos|do servico)\s*(?:\(R\$\))?\s*[:\-]?\s*(?:R\$\s*)?([\d.]+,\d{2})',
        'net_value': r'Valor liquido(?: da NFS-e)?\s*(?:\(R\$\))?\s*[:\-]?\s*(?:R\$\s*)?([\d.]+,\d{2})',
    }
    for field, pattern in patterns.items():
        match = re.search(pattern, plain, re.I)
        if match:
            try:
                raw = match.group(1)
                item[field] = money(raw) if field in MONEY_FIELDS else iso_date(raw) if field in ('issued_at', 'competence') else raw
            except ValueError:
                pass
    keys = set(re.findall(r'(?<!\d)\d{50}(?!\d)', text))
    if len(keys) > 1:
        raise ValueError('PDF contém várias chaves de NFS-e. Separe as notas em arquivos individuais.')
    if keys:
        item['access_key'] = keys.pop()
    return [item]


def classify(item, client_document):
    if item['provider_document'] == client_document:
        return 'Emitida'
    if item['recipient_document'] == client_document:
        return 'Recebida'
    return 'A conferir'


def identity(item):
    if item['access_key']:
        return 'key:' + item['access_key']
    if item['source'].startswith('XML'):
        return 'xml:' + '|'.join(str(item[k] or '') for k in ('provider_document', 'municipality', 'number', 'issued_at', 'verification_code'))
    return ''
