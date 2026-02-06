import os
from datetime import datetime

import pandas as pd
from tqdm import tqdm
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


def gerar_relatorio_pdf(
    df,
    vendas_por_vendedor,
    vendas_por_produto,
    vendas_por_pagamento,
    total_geral,
    caminho_grafico="grafico_vendas_produto.png",
    nome_arquivo="relatorio_vendas.pdf",
):
    largura, altura = A4
    c = canvas.Canvas(nome_arquivo, pagesize=A4)

    # Cabeçalho
    c.setFont("Helvetica-Bold", 20)
    c.drawString(50, altura - 60, "Relatório de Vendas")

    c.setFont("Helvetica", 11)
    c.drawString(50, altura - 80, f"Data de geração: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    # Indicadores principais
    y = altura - 120
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, y, "Resumo geral")

    y -= 20
    c.setFont("Helvetica", 11)
    c.drawString(50, y, f"Total geral de vendas: R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    # Ticket médio (se houver vendas)
    if len(df) > 0:
        ticket_medio = total_geral / len(df)
        y -= 15
        c.drawString(50, y, f"Ticket médio por venda: R$ {ticket_medio:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    # Destaques
    if not vendas_por_produto.empty:
        produto_top = vendas_por_produto.idxmax()
        valor_top_produto = vendas_por_produto.max()
        participacao_produto = (valor_top_produto / total_geral) * 100 if total_geral else 0

        y -= 25
        c.setFont("Helvetica-Bold", 14)
        c.drawString(50, y, "Destaques")

        y -= 20
        c.setFont("Helvetica", 11)
        c.drawString(
            50,
            y,
            f"Produto destaque: {produto_top} (R$ {valor_top_produto:,.2f}, {participacao_produto:.1f}% do total)"
            .replace(",", "X").replace(".", ",").replace("X", "."),
        )

    if not vendas_por_vendedor.empty:
        vendedor_top = vendas_por_vendedor.idxmax()
        valor_top_vendedor = vendas_por_vendedor.max()
        media_vendedores = vendas_por_vendedor.mean()
        multiplicador = valor_top_vendedor / media_vendedores if media_vendedores else 0

        y -= 15
        c.drawString(
            50,
            y,
            f"Melhor vendedor: {vendedor_top} (R$ {valor_top_vendedor:,.2f}, {multiplicador:.1f}x a média dos vendedores)"
            .replace(",", "X").replace(".", ",").replace("X", "."),
        )

    if not vendas_por_pagamento.empty:
        forma_top = vendas_por_pagamento.idxmax()
        valor_top_pagamento = vendas_por_pagamento.max()
        participacao_pagamento = (valor_top_pagamento / total_geral) * 100 if total_geral else 0

        y -= 15
        c.drawString(
            50,
            y,
            f"Forma de pagamento mais usada: {forma_top} (R$ {valor_top_pagamento:,.2f}, {participacao_pagamento:.1f}% do total)"
            .replace(",", "X").replace(".", ",").replace("X", "."),
        )

    # Pequenas tabelas de Top 5
    y -= 30
    c.setFont("Helvetica-Bold", 13)
    c.drawString(50, y, "Top 5 produtos (por faturamento)")
    c.setFont("Helvetica", 10)
    y -= 15
    for produto, valor in vendas_por_produto.head(5).items():
        if y < 120:
            break
        c.drawString(
            60,
            y,
            f"- {produto}: R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
        )
        y -= 12

    y -= 10
    c.setFont("Helvetica-Bold", 13)
    c.drawString(320, altura - 150, "Top 5 vendedores")
    c.setFont("Helvetica", 10)
    y_vend = altura - 165
    for vendedor, valor in vendas_por_vendedor.head(5).items():
        if y_vend < 120:
            break
        c.drawString(
            320,
            y_vend,
            f"- {vendedor}: R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
        )
        y_vend -= 12

    # Inserir gráfico, se existir
    if os.path.exists(caminho_grafico):
        c.showPage()
        largura, altura = A4
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, altura - 60, "Gráfico - Total de Vendas por Produto")

        try:
            img = ImageReader(caminho_grafico)
            img_largura = largura - 100
            img_altura = altura - 150
            c.drawImage(
                img,
                50,
                80,
                width=img_largura,
                height=img_altura,
                preserveAspectRatio=True,
                anchor="c",
            )
        except Exception:
            c.setFont("Helvetica", 12)
            c.drawString(50, altura - 90, "Não foi possível carregar o gráfico para o PDF.")

    c.showPage()
    c.save()


tqdm.pandas(desc="Processando vendas")

df = pd.read_csv("vendas_loja.csv")

df.dropna(subset=["Produto", "Quantidade", "Valor Unitário"], inplace=True)
df["Total"] = df["Quantidade"] * df["Valor Unitário"]
total_geral = df["Total"].sum()
vendas_por_vendedor = df.groupby("Vendedor")["Total"].sum().sort_values(ascending=False)
vendas_por_produto = df.groupby("Produto")["Total"].sum().sort_values(ascending=False)
vendas_por_pagamento = df.groupby("Forma de Pagamento")["Total"].sum().sort_values(ascending=False)

with pd.ExcelWriter("analise_vendas.xlsx") as writer:
    df.to_excel(writer, sheet_name="Vendas Detalhadas", index=False)
    vendas_por_vendedor.to_excel(writer, sheet_name="Por vendedor")
    vendas_por_produto.to_excel(writer, sheet_name="Por produto")
    vendas_por_pagamento.to_excel(writer, sheet_name="Por pagamento")

plt.figure(figsize=(8, 5))
vendas_por_produto.plot(kind="bar", color="skyblue")
plt.title("Total de Vendas por Produto")
plt.ylabel("Valor Total (R$)")
plt.xlabel("Produto")
plt.tight_layout()
plt.savefig("grafico_vendas_produto.png")
plt.show()

print("\nCalculando totais individuais...")
df["Total"].progress_apply(lambda x: x)

print("\n===== RESUMO GERAL =====")
print(f"Total geral de vendas: R$ {total_geral:,.2f}")
print("\nPor vendedor:")
print(vendas_por_vendedor)
print("\nPor produto:")
print(vendas_por_produto)
print("\nPor pagamento:")
print(vendas_por_pagamento)

gerar_relatorio_pdf(
    df=df,
    vendas_por_vendedor=vendas_por_vendedor,
    vendas_por_produto=vendas_por_produto,
    vendas_por_pagamento=vendas_por_pagamento,
    total_geral=total_geral,
)