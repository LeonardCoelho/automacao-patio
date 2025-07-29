# 🚛 Automação de Atualização - Controle de Pátio

Este projeto automatiza a atualização de uma planilha de **Controle de Pátio**, colando automaticamente dados de chegada na aba correta de um arquivo final. Isso substitui o trabalho manual que fazemos no dia a dia na logística, economizando tempo e evitando erros.

---

## 📌 Objetivo

Copiar automaticamente as datas de chegada de uma base extraída do sistema para o controle principal, onde fórmulas como `PROCV` fazem o resto da mágica.

---

## 🛠️ Tecnologias

- Python 3.x
- [xlwings](https://docs.xlwings.org/)
- pandas
- glob / os / time

---

## 📁 Estrutura do Projeto

```
automacao-patio/
├── 📁 data/                  # Contém exemplos de planilhas base (.xlsx)
│   └── base_exemplo.xlsx
├── 📁 images/                # Prints e imagens para README ou documentação
│   └── print_final.jpg
├── 📁 src/                   # Código-fonte do projeto
│   └── atualizar_planilha.py
├── README.md
└── requirements.txt         # Lista de dependências
```

---

## ▶️ Como usar

1. Baixe a base do sistema e salve na pasta `Downloads` do seu usuário
2. Verifique se o arquivo de destino está **fechado**
3. Rode o script:

```bash
python src/atualizar_planilha.py
```

4. O script vai:
   - Encontrar a base mais recente no `Downloads`
   - Achar o arquivo de destino na pasta da rede da Arcor
   - Substituir os dados da aba `BASE` com a nova base
   - Abrir a planilha final atualizada

---

## ✅ Benefícios

- Evita copiar e colar manualmente
- Garante consistência nas atualizações
- Dá pra rodar o script com um clique e seguir o baile no operacional

---

## 📋 Exemplo visual

![Exemplo da planilha atualizada](./images/Print_final.jpg)

---

## 💼 Autor

**Leonardo Coelho**  
[GitHub](https://github.com/LeonardCoelho) · [LinkedIn](https://www.linkedin.com/in/leonardocoelho/)  
Campinas/SP

---

> Feito pra resolver dor real do dia a dia com código na veia. 🚀
