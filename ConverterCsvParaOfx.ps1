param (
  [Parameter(Mandatory)]$ArquivoEntradaCsv,
  $ArquivoSaidaOfxComExtensao = '.\Resultado.ofx'
)

# Importa o arquivo CSV para memória
$transactions = Import-Csv -Path $ArquivoEntradaCsv

# Preenche os dados de Header do OFX
$header = @"
OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE
"@

# Cria o corpo do arquivo OFX
$body = @"
<OFX>
  <SIGNONMSGSRSV1>
    <SONRS>
      <STATUS>
        <CODE>0</CODE>
        <SEVERITY>INFO</SEVERITY>
      </STATUS>
      <DTSERVER>$(Get-Date -Format yyyyMMddHHmmss)</DTSERVER>
      <LANGUAGE>POR</LANGUAGE>
    </SONRS>
  </SIGNONMSGSRSV1>
  <BANKMSGSRSV1>
    <STMTTRNRS>
      <TRNUID>1001</TRNUID>
      <STATUS>
        <CODE>0</CODE>
        <SEVERITY>INFO</SEVERITY>
      </STATUS>
      <STMTRS>
        <CURDEF>BRL</CURDEF>
        <BANKACCTFROM>
          <BANKID>000000000</BANKID>
          <ACCTID>$(New-Guid)</ACCTID>
          <ACCTTYPE>CHECKING</ACCTTYPE>
        </BANKACCTFROM>
        <BANKTRANLIST>
          <DTSTART>$(($transactions | Sort-Object -Property date)[0].Date -as [DateTime]).ToString("yyyyMMdd")</DTSTART>
          <DTEND>$(($transactions | Sort-Object -Property date -Descending)[0].date -as [DateTime]).ToString("yyyyMMdd")</DTEND>
          $(foreach ($transaction in $transactions) {
              "<STMTTRN>"
              "<TRNTYPE>$(($transaction.category -replace "&", "&amp;"))</TRNTYPE>" #Categoria da transação
              "<DTPOSTED>$(($transaction.date -as [DateTime]).ToString("yyyyMMdd"))</DTPOSTED>"
              "<TRNAMT>$(("{0:F2}" -f ($transaction.amount -replace '"', '' -replace ",", ".") ))</TRNAMT>" #Valor da transação
              "<FITID>$(New-Guid)</FITID>" #GUID Único da transação
              "<NAME>$(($transaction.title -replace "&", "&amp;"))</NAME>" #Nome da transação
              "</STMTTRN>"
          })
        </BANKTRANLIST>
        <LEDGERBAL>
          <BALAMT>0.00</BALAMT>
          <DTASOF>$(Get-Date -Format yyyyMMddHHmmss)</DTASOF>
        </LEDGERBAL>
      </STMTRS>
    </STMTTRNRS>
  </BANKMSGSRSV1>
</OFX>
"@

# Escreve o arquivo FX Resultante
$header + $body | Out-File -Encoding UTF8 -FilePath $ArquivoSaidaOfxComExtensao