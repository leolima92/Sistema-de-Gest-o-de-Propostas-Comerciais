import os
import sqlite3
from datetime import datetime
from typing import Dict

from ..models import GestorPropostas, Cliente, Proposta, ItemProposta

BASE_DIR = os.path.dirname(__file__)


class StorageManager:
    DB_PATH = os.path.join(BASE_DIR, "gestor_propostas.db")

    @classmethod
    def _get_conn(cls):
        return sqlite3.connect(cls.DB_PATH)

    @classmethod
    def init_db(cls):
        conn = cls._get_conn()
        cur = conn.cursor()

        # Clientes
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY,
                nome TEXT NOT NULL,
                documento TEXT,
                contato TEXT
            )
            """
        )

        # Propostas
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS propostas (
                id INTEGER PRIMARY KEY,
                cliente_id INTEGER NOT NULL,
                titulo TEXT NOT NULL,
                data_criacao TEXT NOT NULL,
                status TEXT NOT NULL,
                validade TEXT,
                responsavel TEXT,
                condicoes_pagamento TEXT,
                tipo_desconto TEXT,
                desconto_percentual REAL,
                desconto_valor REAL,
                FOREIGN KEY (cliente_id) REFERENCES clientes(id)
            )
            """
        )

        # Itens
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS itens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                proposta_id INTEGER NOT NULL,
                descricao TEXT NOT NULL,
                quantidade INTEGER NOT NULL,
                valor_unitario REAL NOT NULL,
                FOREIGN KEY (proposta_id) REFERENCES propostas(id)
            )
            """
        )

        conn.commit()
        conn.close()

    @classmethod
    def salvar_ou_atualizar_cliente(cls, cliente: Cliente):
        conn = cls._get_conn()
        cur = conn.cursor()

        cur.execute("SELECT 1 FROM clientes WHERE id = ?", (cliente.id,))
        existe = cur.fetchone() is not None

        if not existe:
            cur.execute(
                """
                INSERT INTO clientes (id, nome, documento, contato)
                VALUES (?, ?, ?, ?)
                """,
                (cliente.id, cliente.nome, cliente.documento, cliente.contato),
            )
        else:
            cur.execute(
                """
                UPDATE clientes
                   SET nome = ?,
                       documento = ?,
                       contato = ?
                 WHERE id = ?
                """,
                (cliente.nome, cliente.documento, cliente.contato, cliente.id),
            )

        conn.commit()
        conn.close()

    @classmethod
    def deletar_cliente(cls, cliente_id: int):
        conn = cls._get_conn()
        cur = conn.cursor()

        cur.execute("DELETE FROM clientes WHERE id = ?", (cliente_id,))

        conn.commit()
        conn.close()

    @classmethod
    def salvar_ou_atualizar_proposta(cls, proposta: Proposta):
        conn = cls._get_conn()
        cur = conn.cursor()

        cur.execute("SELECT 1 FROM propostas WHERE id = ?", (proposta.id,))
        existe = cur.fetchone() is not None

        data_criacao_str = proposta.data_criacao.strftime("%Y-%m-%d %H:%M:%S")
        validade_str = proposta.validade.strftime("%Y-%m-%d") if proposta.validade else None

        if not existe:
            cur.execute(
                """
                INSERT INTO propostas (
                    id, cliente_id, titulo, data_criacao, status,
                    validade, responsavel, condicoes_pagamento,
                    tipo_desconto, desconto_percentual, desconto_valor
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    proposta.id,
                    proposta.cliente.id,
                    proposta.titulo,
                    data_criacao_str,
                    proposta.status,
                    validade_str,
                    proposta.responsavel,
                    proposta.condicoes_pagamento,
                    proposta.tipo_desconto,
                    proposta.desconto_percentual,
                    proposta.desconto_valor,
                ),
            )
        else:
            cur.execute(
                """
                UPDATE propostas
                   SET cliente_id = ?,
                       titulo = ?,
                       data_criacao = ?,
                       status = ?,
                       validade = ?,
                       responsavel = ?,
                       condicoes_pagamento = ?,
                       tipo_desconto = ?,
                       desconto_percentual = ?,
                       desconto_valor = ?
                 WHERE id = ?
                """,
                (
                    proposta.cliente.id,
                    proposta.titulo,
                    data_criacao_str,
                    proposta.status,
                    validade_str,
                    proposta.responsavel,
                    proposta.condicoes_pagamento,
                    proposta.tipo_desconto,
                    proposta.desconto_percentual,
                    proposta.desconto_valor,
                    proposta.id,
                ),
            )

        conn.commit()
        conn.close()

    @classmethod
    def deletar_proposta(cls, proposta_id: int):
        conn = cls._get_conn()
        cur = conn.cursor()

        cur.execute("DELETE FROM itens WHERE proposta_id = ?", (proposta_id,))
        cur.execute("DELETE FROM propostas WHERE id = ?", (proposta_id,))

        conn.commit()
        conn.close()

    # =========================================================
    #   ITENS
    # =========================================================
    @classmethod
    def sincronizar_itens_proposta(cls, proposta: Proposta):
        conn = cls._get_conn()
        cur = conn.cursor()

        cur.execute("DELETE FROM itens WHERE proposta_id = ?", (proposta.id,))

        for item in proposta.itens:
            cur.execute(
                """
                INSERT INTO itens (proposta_id, descricao, quantidade, valor_unitario)
                VALUES (?, ?, ?, ?)
                """,
                (proposta.id, item.descricao, item.quantidade, item.valor_unitario),
            )

        conn.commit()
        conn.close()

    @classmethod
    def carregar_tudo(cls, gestor: GestorPropostas):
        gestor.clientes.clear()
        gestor.propostas.clear()

        conn = cls._get_conn()
        cur = conn.cursor()

        # ---- Clientes
        cur.execute("SELECT id, nome, documento, contato FROM clientes ORDER BY id")
        rows_clientes = cur.fetchall()

        mapa_clientes: Dict[int, Cliente] = {}
        max_cliente_id = 0

        for row in rows_clientes:
            cli_id, nome, documento, contato = row
            cliente = Cliente(nome, documento or "", contato or "")
            cliente.id = cli_id
            mapa_clientes[cli_id] = cliente
            gestor.clientes.append(cliente)
            max_cliente_id = max(max_cliente_id, cli_id)

        if max_cliente_id > 0:
            Cliente._contador_id = max_cliente_id + 1

        cur.execute(
            """
            SELECT
                id, cliente_id, titulo, data_criacao, status,
                validade, responsavel, condicoes_pagamento,
                tipo_desconto, desconto_percentual, desconto_valor
            FROM propostas
            ORDER BY id
            """
        )
        rows_propostas = cur.fetchall()

        mapa_propostas: Dict[int, Proposta] = {}
        max_proposta_id = 0

        for row in rows_propostas:
            (
                p_id,
                cliente_id,
                titulo,
                data_criacao_str,
                status,
                validade_str,
                responsavel,
                condicoes_pagamento,
                tipo_desconto,
                desconto_percentual,
                desconto_valor,
            ) = row

            cliente = mapa_clientes.get(cliente_id)
            if not cliente:
                continue

            prop = Proposta(
                cliente=cliente,
                titulo=titulo,
                validade=None,
                responsavel=responsavel or "",
                condicoes_pagamento=condicoes_pagamento or "",
            )

            prop.id = p_id
            try:
                prop.data_criacao = datetime.strptime(
                    data_criacao_str, "%Y-%m-%d %H:%M:%S"
                )
            except Exception:
                prop.data_criacao = datetime.now()

            prop.status = status
            if validade_str:
                try:
                    prop.validade = datetime.strptime(
                        validade_str, "%Y-%m-%d"
                    ).date()
                except Exception:
                    prop.validade = None

            prop.tipo_desconto = tipo_desconto
            prop.desconto_percentual = desconto_percentual or 0.0
            prop.desconto_valor = desconto_valor or 0.0

            gestor.propostas.append(prop)
            mapa_propostas[p_id] = prop
            max_proposta_id = max(max_proposta_id, p_id)

        if max_proposta_id > 0:
            Proposta._contador_id = max_proposta_id + 1

        cur.execute(
            """
            SELECT proposta_id, descricao, quantidade, valor_unitario
            FROM itens
            ORDER BY id
            """
        )
        rows_itens = cur.fetchall()

        for row in rows_itens:
            proposta_id, descricao, quantidade, valor_unitario = row
            prop = mapa_propostas.get(proposta_id)
            if not prop:
                continue
            item = ItemProposta(
                descricao=descricao,
                quantidade=int(quantidade),
                valor_unitario=float(valor_unitario),
            )
            prop.itens.append(item)

        conn.close()
