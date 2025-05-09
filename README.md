# gerador-oficios
Gera ofícios Word a partir de um modelo e uma planilha Excel.

## Instruções de uso:
  ### 1. Crie (opcional) e ative um ambiente virtual
    python -m venv .venv
    source .venv/bin/activate  #Linux/Mac
    .venv\Scripts\activate      # Windows
  ### 2. Instale dependências
    pip install -r requirements.txt

## Arquivos esperados (mesma pasta do script):
  modelo_oficio.docx
  dados_oficios.xlsx

  ### A plainlha dados_oficios.xlsx deve conter os seguintes campos:
    | n  | dia | mes  | sexo | nome          | cargo                        |

  #### Exemplo de preenchimento:

    | n  | dia | mes  | sexo | nome          | cargo                        |
    | -- | --- | ---- | ---- | ------------- | ---------------------------- |
    | 23 | 10  | maio | M    | João da Silva | Diretor-Geral da ANTT        |
    | 24 | 10  | maio | F    | Maria Soares  | Secretária de Infraestrutura |
## Saída:
    ./oficios_gerados/Oficio_###_PrimeiroNome.docx