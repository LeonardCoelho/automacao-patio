# 🚛 Automação de Atualização - Controle de Pátio

Este script em Python automatiza a atualização de uma planilha de **Controle de Pátio**, substituindo um processo manual de copiar e colar dados de chegada. Ele lê uma base baixada da web e cola automaticamente os dados atualizados na aba correta do arquivo principal. Isso economiza tempo, evita erros e agiliza o fluxo do time logístico.

---

## 📌 Objetivo

Substituir o trabalho manual de:
- Abrir planilha baixada com datas de chegada
- Copiar as informações
- Colar na aba `BASE` do controle geral, que usa `PROCV` para preencher os dados dos romaneios

---

## 🛠️ Tecnologias utilizadas

- Python 3.x
- [xlwings](https://docs.xlwings.org/en/stable/) (automação Excel via Python)
- pandas
- glob / os / time (controle de arquivos)

---

## 📂 Estrutura do Projeto

```
automacao-patio/
├── automacao_patio.py       # Script principal
├── README.md
├── requirements.txt         # (opcional)
```

---

## ▶️ Como usar

1. Baixe manualmente a base do sistema (relatório do TMS) e salve na pasta `Downloads`
2. Garanta que o arquivo de destino (controle principal) esteja fechado
3. Rode o script com Python:

```bash
python automacao_patio.py
```

4. O script vai:
   - Encontrar automaticamente a base mais recente no `Downloads`
   - Encontrar o arquivo de controle de pátio na pasta de rede
   - Substituir os dados da aba `BASE` com a nova base
   - Abrir a planilha final atualizada para revisão

---

## 📍 Requisitos

- Python instalado no PC
- Permissão de leitura/gravação na pasta de rede da Arcor
- Excel instalado (necessário pro `xlwings` funcionar)

---

## ⚙️ Parametrizações do Script

- Nome do arquivo base: `"Controle de Patio x data ingreso do pedido"`
- Nome da aba a ser atualizada: `"BASE"`
- Pasta destino: `"Arcor\GC_CPS_CDs_Controle_de_Patio - Documentos\CONTROLE DE PÁTIO - 2025"`
- O script localiza automaticamente os arquivos mais recentes com base no padrão do nome

---

## 💡 Benefícios

✅ Agilidade no processo de atualização  
✅ Redução de erro humano  
✅ Mais tempo pra focar em análise e controle logístico  

---

## 👤 Autor

**Leonardo Coelho**  
📍 Campinas/SP  
[LinkedIn](https://www.linkedin.com/in/leonardocoelho/)  

> Projeto feito com amor (e café) pra facilitar o trampo diário na logística 🚛📊
