from datetime import datetime
from typing import List, Optional

class Cliente:
    _contador_id = 1

    def __init__(self, nome: str, documento: str = "", contato: str = ""):
        self.id = Cliente._contador_id
        Cliente._contador_id += 1

        self.nome = nome
        self.documento = documento
        self.contato = contato

    def __str__(self) -> str:
        doc = f" | Doc: {self.documento}" if self.documento else ""
        contato = f" | Contato: {self.contato}" if self.contato else ""
        return f"({self.id}) {self.nome}{doc}{contato}"


class ItemProposta:
    def __init__(self, descricao: str, quantidade: int, valor_unitario: float):
        self.descricao = descricao
        self.quantidade = quantidade
        self.valor_unitario = valor_unitario

    @property
    def total(self) -> float:
        return self.quantidade * self.valor_unitario

    def __str__(self) -> str:
        return (
            f"{self.descricao} | Qtd: {self.quantidade} | "
            f"Unit: R$ {self.valor_unitario:.2f} | Total: R$ {self.total:.2f}"
        )


class Proposta:
    _contador_id = 1
    STATUS_VALIDOS = ["rascunho", "enviada", "aceita", "recusada", "cancelada"]

    def __init__(
        self,
        cliente: Cliente,
        titulo: str = "",
        validade=None,       
        responsavel: str = "",
        condicoes_pagamento: str = "",
    ):
        self.id = Proposta._contador_id
        Proposta._contador_id += 1

        self.cliente = cliente
        self.titulo = titulo or f"Proposta {self.id}"
        self.data_criacao = datetime.now()
        self.status = "rascunho"
        self.itens: List[ItemProposta] = []
        self.validade = validade         
        self.responsavel = responsavel
        self.condicoes_pagamento = condicoes_pagamento

        self.tipo_desconto = None  
        self.desconto_percentual = 0.0
        self.desconto_valor = 0.0

    def adicionar_item(self, item: ItemProposta):
        self.itens.append(item)

    def calcular_subtotal(self) -> float:
        return sum(item.total for item in self.itens)

    def definir_desconto_percentual(self, percentual: float):
        self.tipo_desconto = "%"
        self.desconto_percentual = max(0.0, percentual)
        self.desconto_valor = 0.0

    def definir_desconto_valor(self, valor: float):
        self.tipo_desconto = "R"
        self.desconto_valor = max(0.0, valor)
        self.desconto_percentual = 0.0

    def calcular_desconto(self) -> float:
        subtotal = self.calcular_subtotal()
        if self.tipo_desconto == "%":
            return subtotal * (self.desconto_percentual / 100.0)
        elif self.tipo_desconto == "R":
            return self.desconto_valor
        return 0.0

    def calcular_total(self) -> float:
        subtotal = self.calcular_subtotal()
        desconto = self.calcular_desconto()
        return max(0.0, subtotal - desconto)

    def alterar_status(self, novo_status: str):
        novo_status = novo_status.lower()
        if novo_status not in Proposta.STATUS_VALIDOS:
            raise ValueError(f"Status inválido: {novo_status}")
        self.status = novo_status

    def __str__(self) -> str:
        subtotal = self.calcular_subtotal()
        total = self.calcular_total()
        return (
            f"#{self.id} - {self.titulo} | Cliente: {self.cliente.nome} | "
            f"Status: {self.status} | Itens: {len(self.itens)} | "
            f"Subtotal: R$ {subtotal:.2f} | Total: R$ {total:.2f}"
        )


class GestorPropostas:
    def __init__(self):
        self.clientes: List[Cliente] = []
        self.propostas: List[Proposta] = []

    def criar_cliente(self, nome: str, documento: str = "", contato: str = "") -> Cliente:
        cliente = Cliente(nome, documento, contato)
        self.clientes.append(cliente)
        return cliente

    def listar_clientes(self) -> List[Cliente]:
        return self.clientes

    def obter_cliente_por_indice(self, indice: int) -> Optional[Cliente]:
        if 0 <= indice < len(self.clientes):
            return self.clientes[indice]
        return None

    # ---- Propostas ---- #

    def criar_proposta(
        self,
        cliente: Cliente,
        titulo: str = "",
        validade=None,
        responsavel: str = "",
        condicoes_pagamento: str = "",
    ) -> Proposta:
        proposta = Proposta(
            cliente,
            titulo,
            validade=validade,
            responsavel=responsavel,
            condicoes_pagamento=condicoes_pagamento,
        )
        self.propostas.append(proposta)
        return proposta

    def listar_propostas(self) -> List[Proposta]:
        return self.propostas

    def obter_proposta_por_indice(self, indice: int) -> Optional[Proposta]:
        if 0 <= indice < len(self.propostas):
            return self.propostas[indice]
        return None
