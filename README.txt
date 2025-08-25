
# Calculadora de Emissões e Redução de CO₂ com Hidrogênio

Este pacote contém um aplicativo **Streamlit** que lê a planilha `Calculos.xlsx` e calcula:
- Emissão do combustível original (kg CO₂);
- Quantidade de H₂ necessária para substituir;
- Emissões do H₂ por rota de produção (Eletrólise, Biomassa, SMR, SMR+CCS);
- Redução de CO₂ em cada cenário.

## Como executar
1) Instale as dependências (de preferência em um ambiente virtual):
   ```bash
   pip install streamlit pandas openpyxl
   ```

2) Garanta que o arquivo **Calculos.xlsx** esteja na mesma pasta de `app.py`. 
   - Você pode usar o seu arquivo real ou o exemplo incluído aqui.

3) Rode o aplicativo:
   ```bash
   streamlit run app.py
   ```

4) O navegador abrirá a interface. Se preferir, você pode também **fazer upload** de outra planilha diretamente pela interface.

## Compatibilidade de formatos de planilha
O app tenta automaticamente detectar:
- **Formato “tidy” (arrumado)** com colunas: `Combustível`, `Fator_CO2 (kg/kg)` e `H2_equivalente (kg/kg)`.
- **Formato “matricial” (como no seu Excel)**, usando por padrão:
  - Linha 0 = quantidade base do combustível (ex.: 1 kg);
  - Linha 2 = fator “kg H2 / kg combustível” (H2 equivalente);
  - Linha 4 = emissões de referência (kg CO₂), normalmente com cabeçalhos como “Emissão Gás Natural”.
  - Você pode ajustar esses índices no painel **⚙️ Opções avançadas** do app.

## Observações
- Os fatores de emissão do H₂ (Eletrólise, Biomassa, SMR, SMR+CCS) são **editáveis** na interface.
- Exporte os resultados em CSV pelo botão de download.
