import os
import tempfile
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

class ExcelReportGenerator:
    @classmethod
    def gerar_excel(cls, gestor, caminho: str | None = None) -> str:
        if not caminho:
            fd, caminho = tempfile.mkstemp(
                suffix=".xlsx", prefix="dealflow_propostas_"
            )
            os.close(fd)

        wb = Workbook()
        ws = wb.active
        ws.title = "Propostas"

        # Cabeçalho
        headers = [
            "ID Proposta",
            "Título",
            "Cliente",
            "Documento cliente",
            "Contato cliente",
            "Status",
            "Data criação",
            "Responsável",
            "Validade",
            "Condições de pagamento",
            "Subtotal (R$)",
            "Desconto (R$)",
            "Total (R$)",
        ]
        ws.append(headers)

        # Estilo do cabeçalho
        header_font = Font(bold=True)
        for cell in ws[1]:
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        # Linhas de dados
        for p in gestor.listar_propostas():
            subtotal = p.calcular_subtotal()
            total = p.calcular_total()
            desconto = subtotal - total

            data_criacao_str = (
                p.data_criacao.strftime("%d/%m/%Y %H:%M")
                if p.data_criacao
                else ""
            )
            validade_str = (
                p.validade.strftime("%d/%m/%Y") if p.validade else ""
            )

            ws.append(
                [
                    p.id,
                    p.titulo,
                    p.cliente.nome,
                    p.cliente.documento,
                    p.cliente.contato,
                    p.status,
                    data_criacao_str,
                    p.responsavel or "",
                    validade_str,
                    p.condicoes_pagamento or "",
                    float(subtotal),
                    float(desconto),
                    float(total),
                ]
            )


        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    value = str(cell.value) if cell.value is not None else ""
                    if len(value) > max_len:
                        max_len = len(value)
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = max(10, min(max_len + 2, 50))

        wb.save(caminho)
        return caminho
