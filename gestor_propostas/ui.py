from datetime import datetime
from typing import Optional

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from .models import GestorPropostas, Cliente, Proposta, ItemProposta
from .excel_report import ExcelReportGenerator


class App(tk.Tk):
    def __init__(self, gestor: GestorPropostas):
        super().__init__()
        self.title("Sistema de Gestão de Propostas Comerciais")
        self.geometry("1000x550")

        self.gestor = gestor
        self.filtro_status_var = tk.StringVar(value="Todos")
        self.busca_var = tk.StringVar(value="")  # campo de busca

        self._criar_widgets()


    def _criar_widgets(self):
        # Frame de clientes
        frame_clientes = ttk.LabelFrame(self, text="Clientes")
        frame_clientes.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        self.listbox_clientes = tk.Listbox(frame_clientes, height=15, width=35)
        self.listbox_clientes.pack(side=tk.TOP, fill=tk.BOTH, padx=5, pady=5)

        btn_novo_cliente = ttk.Button(
            frame_clientes, text="Novo Cliente", command=self.janela_novo_cliente
        )
        btn_novo_cliente.pack(side=tk.TOP, fill=tk.X, padx=5, pady=2)

        # Frame de propostas
        frame_propostas = ttk.LabelFrame(self, text="Propostas")
        frame_propostas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Linha de filtro + busca
        frame_filtro = ttk.Frame(frame_propostas)
        frame_filtro.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        ttk.Label(frame_filtro, text="Filtro status:").pack(side=tk.LEFT)
        self.combo_filtro_status = ttk.Combobox(
            frame_filtro,
            values=["Todos"] + Proposta.STATUS_VALIDOS,
            textvariable=self.filtro_status_var,
            state="readonly",
            width=15,
        )
        self.combo_filtro_status.pack(side=tk.LEFT, padx=5)
        self.combo_filtro_status.bind("<<ComboboxSelected>>", lambda e: self.atualizar_listas())

        # Campo de busca
        ttk.Label(frame_filtro, text="Buscar (cliente/título):").pack(side=tk.LEFT, padx=(15, 0))
        entry_busca = ttk.Entry(frame_filtro, textvariable=self.busca_var, width=30)
        entry_busca.pack(side=tk.LEFT, padx=5)

        # Enter também dispara a busca
        entry_busca.bind("<Return>", lambda e: self.atualizar_listas())

        ttk.Button(frame_filtro, text="Buscar", command=self.atualizar_listas)\
            .pack(side=tk.LEFT, padx=2)
        ttk.Button(frame_filtro, text="Limpar", command=self.limpar_busca)\
            .pack(side=tk.LEFT, padx=2)

        # Lista de propostas
        self.listbox_propostas = tk.Listbox(frame_propostas, height=10)
        self.listbox_propostas.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.listbox_propostas.bind("<<ListboxSelect>>", lambda e: self.atualizar_itens_treeview())

        # Tabela de itens da proposta selecionada
        ttk.Label(frame_propostas, text="Itens da proposta selecionada:")\
            .pack(anchor="w", padx=5, pady=(5, 0))

        colunas = ("descricao", "quantidade", "unitario", "total")
        self.tree_itens = ttk.Treeview(
            frame_propostas,
            columns=colunas,
            show="headings",
            height=7,
        )
        self.tree_itens.heading("descricao", text="Descrição")
        self.tree_itens.heading("quantidade", text="Qtd")
        self.tree_itens.heading("unitario", text="Valor Unitário")
        self.tree_itens.heading("total", text="Total")

        self.tree_itens.column("descricao", width=300)
        self.tree_itens.column("quantidade", width=60, anchor="e")
        self.tree_itens.column("unitario", width=100, anchor="e")
        self.tree_itens.column("total", width=100, anchor="e")

        self.tree_itens.pack(side=tk.TOP, fill=tk.BOTH, expand=False, padx=5, pady=5)

        # Botões de ações de proposta
        frame_botoes = ttk.Frame(frame_propostas)
        frame_botoes.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)

        ttk.Button(frame_botoes, text="Nova Proposta", command=self.janela_nova_proposta)\
            .grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        ttk.Button(frame_botoes, text="Adicionar Itens", command=self.janela_adicionar_itens)\
            .grid(row=0, column=1, padx=2, pady=2, sticky="ew")
        ttk.Button(frame_botoes, text="Definir Desconto", command=self.janela_desconto)\
            .grid(row=0, column=2, padx=2, pady=2, sticky="ew")
        ttk.Button(frame_botoes, text="Alterar Status", command=self.janela_status)\
            .grid(row=0, column=3, padx=2, pady=2, sticky="ew")
        ttk.Button(frame_botoes, text="Duplicar Proposta", command=self.acao_duplicar_proposta)\
            .grid(row=0, column=4, padx=2, pady=2, sticky="ew")
        ttk.Button(frame_botoes, text="Gerar Excel", command=self.acao_gerar_excel)\
            .grid(row=0, column=5, padx=2, pady=2, sticky="ew")

        # Linha de botões de edição
        ttk.Button(frame_botoes, text="Editar Proposta", command=self.janela_editar_proposta)\
            .grid(row=1, column=0, padx=2, pady=2, sticky="ew")
        ttk.Button(frame_botoes, text="Editar Item", command=self.janela_editar_item)\
            .grid(row=1, column=1, padx=2, pady=2, sticky="ew")

        for i in range(6):
            frame_botoes.columnconfigure(i, weight=1)

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Pronto.")
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor="w")
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)


    def _get_propostas_filtradas(self):
        """Retorna a lista de propostas respeitando:
        - filtro de status
        - texto de busca (cliente / título)
        """
        propostas = self.gestor.listar_propostas()

        # filtro por status
        filtro = self.filtro_status_var.get()
        if filtro != "Todos":
            propostas = [p for p in propostas if p.status == filtro]

        # filtro por texto (cliente ou título)
        texto = self.busca_var.get().strip().lower()
        if texto:
            propostas = [
                p for p in propostas
                if texto in p.titulo.lower()
                or texto in p.cliente.nome.lower()
            ]

        return propostas

    def atualizar_listas(self):
        # Clientes
        self.listbox_clientes.delete(0, tk.END)
        for c in self.gestor.listar_clientes():
            self.listbox_clientes.insert(tk.END, str(c))

        # Propostas (aplicando filtros)
        self.listbox_propostas.delete(0, tk.END)
        propostas = self._get_propostas_filtradas()

        for p in propostas:
            self.listbox_propostas.insert(tk.END, str(p))

        # Atualiza itens da proposta selecionada (se houver)
        self.atualizar_itens_treeview()

        # Atualiza status bar com total encontrado
        self.status_var.set(f"{len(propostas)} proposta(s) encontrada(s).")

    def atualizar_itens_treeview(self):
        # Limpa a tabela
        for item in self.tree_itens.get_children():
            self.tree_itens.delete(item)

        proposta = self._obter_proposta_selecionada()
        if not proposta:
            return

        for item in proposta.itens:
            self.tree_itens.insert(
                "",
                tk.END,
                values=(
                    item.descricao,
                    item.quantidade,
                    f"R$ {item.valor_unitario:.2f}",
                    f"R$ {item.total:.2f}",
                ),
            )

    def limpar_busca(self):
        self.busca_var.set("")
        self.atualizar_listas()

    def _obter_cliente_selecionado(self) -> Optional[Cliente]:
        idx = self.listbox_clientes.curselection()
        if not idx:
            return None
        return self.gestor.obter_cliente_por_indice(idx[0])

    def _obter_proposta_selecionada(self) -> Optional[Proposta]:
        idx = self.listbox_propostas.curselection()
        if not idx:
            return None

        propostas = self._get_propostas_filtradas()

        if 0 <= idx[0] < len(propostas):
            return propostas[idx[0]]
        return None


    def janela_novo_cliente(self):
        win = tk.Toplevel(self)
        win.title("Novo Cliente")
        win.geometry("350x200")

        ttk.Label(win, text="Nome:").pack(anchor="w", padx=10, pady=5)
        entry_nome = ttk.Entry(win)
        entry_nome.pack(fill=tk.X, padx=10)

        ttk.Label(win, text="Documento (opcional):").pack(anchor="w", padx=10, pady=5)
        entry_doc = ttk.Entry(win)
        entry_doc.pack(fill=tk.X, padx=10)

        ttk.Label(win, text="Contato (opcional):").pack(anchor="w", padx=10, pady=5)
        entry_contato = ttk.Entry(win)
        entry_contato.pack(fill=tk.X, padx=10)

        def salvar():
            nome = entry_nome.get().strip()
            if not nome:
                messagebox.showwarning("Atenção", "Nome é obrigatório.")
                return
            doc = entry_doc.get().strip()
            contato = entry_contato.get().strip()
            self.gestor.criar_cliente(nome, doc, contato)
            self.atualizar_listas()
            win.destroy()
            self.status_var.set(f"Cliente '{nome}' criado com sucesso.")

        ttk.Button(win, text="Salvar", command=salvar).pack(pady=10)

    def janela_nova_proposta(self):
        if not self.gestor.listar_clientes():
            messagebox.showinfo("Informação", "Cadastre um cliente antes de criar uma proposta.")
            return

        win = tk.Toplevel(self)
        win.title("Nova Proposta")
        win.geometry("420x280")

        ttk.Label(win, text="Cliente:").pack(anchor="w", padx=10, pady=5)
        clientes_nomes = [c.nome for c in self.gestor.listar_clientes()]
        combo_cliente = ttk.Combobox(win, values=clientes_nomes, state="readonly")
        combo_cliente.current(0)
        combo_cliente.pack(fill=tk.X, padx=10)

        ttk.Label(win, text="Título da proposta (opcional):").pack(anchor="w", padx=10, pady=5)
        entry_titulo = ttk.Entry(win)
        entry_titulo.pack(fill=tk.X, padx=10)

        ttk.Label(win, text="Responsável (opcional):").pack(anchor="w", padx=10, pady=5)
        entry_responsavel = ttk.Entry(win)
        entry_responsavel.pack(fill=tk.X, padx=10)

        ttk.Label(win, text="Validade (AAAA-MM-DD) [opcional]:").pack(anchor="w", padx=10, pady=5)
        entry_validade = ttk.Entry(win)
        entry_validade.pack(fill=tk.X, padx=10)

        ttk.Label(win, text="Condições de pagamento (opcional):").pack(anchor="w", padx=10, pady=5)
        entry_cond = ttk.Entry(win)
        entry_cond.pack(fill=tk.X, padx=10)

        def salvar():
            nome_cliente = combo_cliente.get()
            cliente = next((c for c in self.gestor.listar_clientes() if c.nome == nome_cliente), None)
            if not cliente:
                messagebox.showerror("Erro", "Cliente inválido.")
                return

            titulo = entry_titulo.get().strip()
            responsavel = entry_responsavel.get().strip()
            cond_pag = entry_cond.get().strip()

            validade_str = entry_validade.get().strip()
            validade = None
            if validade_str:
                try:
                    validade = datetime.strptime(validade_str, "%Y-%m-%d").date()
                except ValueError:
                    messagebox.showwarning(
                        "Atenção",
                        "Data de validade inválida. Use o formato AAAA-MM-DD (ex: 2025-12-31)."
                    )
                    return

            prop = self.gestor.criar_proposta(
                cliente,
                titulo,
                validade=validade,
                responsavel=responsavel,
                condicoes_pagamento=cond_pag,
            )
            self.atualizar_listas()
            win.destroy()
            self.status_var.set(f"Proposta #{prop.id} criada para {cliente.nome}.")

        ttk.Button(win, text="Criar", command=salvar).pack(pady=10)

    def janela_editar_proposta(self):
        proposta = self._obter_proposta_selecionada()
        if not proposta:
            messagebox.showinfo("Informação", "Selecione uma proposta para editar.")
            return

        win = tk.Toplevel(self)
        win.title(f"Editar Proposta #{proposta.id}")
        win.geometry("420x280")

        ttk.Label(win, text=f"Cliente: {proposta.cliente.nome}")\
            .pack(anchor="w", padx=10, pady=5)

        ttk.Label(win, text="Título da proposta:").pack(anchor="w", padx=10, pady=5)
        entry_titulo = ttk.Entry(win)
        entry_titulo.pack(fill=tk.X, padx=10)
        entry_titulo.insert(0, proposta.titulo)

        ttk.Label(win, text="Responsável (opcional):").pack(anchor="w", padx=10, pady=5)
        entry_responsavel = ttk.Entry(win)
        entry_responsavel.pack(fill=tk.X, padx=10)
        entry_responsavel.insert(0, proposta.responsavel or "")

        ttk.Label(win, text="Validade (AAAA-MM-DD) [opcional]:").pack(anchor="w", padx=10, pady=5)
        entry_validade = ttk.Entry(win)
        entry_validade.pack(fill=tk.X, padx=10)
        if proposta.validade:
            entry_validade.insert(0, proposta.validade.strftime("%Y-%m-%d"))

        ttk.Label(win, text="Condições de pagamento (opcional):").pack(anchor="w", padx=10, pady=5)
        entry_cond = ttk.Entry(win)
        entry_cond.pack(fill=tk.X, padx=10)
        entry_cond.insert(0, proposta.condicoes_pagamento or "")

        def salvar():
            titulo = entry_titulo.get().strip()
            if not titulo:
                messagebox.showwarning("Atenção", "Título não pode ficar vazio.")
                return

            responsavel = entry_responsavel.get().strip()
            cond_pag = entry_cond.get().strip()

            validade_str = entry_validade.get().strip()
            validade = None
            if validade_str:
                try:
                    validade = datetime.strptime(validade_str, "%Y-%m-%d").date()
                except ValueError:
                    messagebox.showwarning(
                        "Atenção",
                        "Data de validade inválida. Use o formato AAAA-MM-DD (ex: 2025-12-31)."
                    )
                    return

            proposta.titulo = titulo
            proposta.responsavel = responsavel
            proposta.validade = validade
            proposta.condicoes_pagamento = cond_pag

            self.atualizar_listas()
            win.destroy()
            self.status_var.set(f"Proposta #{proposta.id} atualizada com sucesso.")

        ttk.Button(win, text="Salvar alterações", command=salvar).pack(pady=10)
        
    def janela_desconto(self):
        proposta = self._obter_proposta_selecionada()
        if not proposta:
            messagebox.showinfo("Informação", "Selecione uma proposta.")
            return

        win = tk.Toplevel(self)
        win.title(f"Desconto - Proposta #{proposta.id}")
        win.geometry("320x200")

        ttk.Label(win, text=f"Proposta #{proposta.id} - {proposta.titulo}")\
            .pack(anchor="w", padx=10, pady=5)

        tipo_var = tk.StringVar(value="nenhum")

        frame_tipos = ttk.Frame(win)
        frame_tipos.pack(padx=10, pady=5, anchor="w")

        ttk.Radiobutton(frame_tipos, text="Nenhum", variable=tipo_var, value="nenhum").pack(anchor="w")
        ttk.Radiobutton(frame_tipos, text="Percentual (%)", variable=tipo_var, value="%").pack(anchor="w")
        ttk.Radiobutton(frame_tipos, text="Valor (R$)", variable=tipo_var, value="R").pack(anchor="w")

        ttk.Label(win, text="Valor:").pack(anchor="w", padx=10, pady=5)
        entry_valor = ttk.Entry(win)
        entry_valor.pack(fill=tk.X, padx=10)

        def aplicar():
            tipo = tipo_var.get()
            if tipo == "nenhum":
                proposta.tipo_desconto = None
                proposta.desconto_percentual = 0.0
                proposta.desconto_valor = 0.0
                self.status_var.set(f"Desconto removido da Proposta #{proposta.id}.")
            else:
                try:
                    valor = float(entry_valor.get().strip().replace(",", "."))
                except ValueError:
                    messagebox.showwarning("Atenção", "Informe um valor numérico.")
                    return

                if tipo == "%":
                    proposta.definir_desconto_percentual(valor)
                    self.status_var.set(
                        f"Desconto de {valor:.2f}% aplicado na Proposta #{proposta.id}."
                    )
                elif tipo == "R":
                    proposta.definir_desconto_valor(valor)
                    self.status_var.set(
                        f"Desconto de R$ {valor:.2f} aplicado na Proposta #{proposta.id}."
                    )

            self.atualizar_listas()
            win.destroy()

        ttk.Button(win, text="Aplicar", command=aplicar).pack(pady=10)

    def janela_status(self):
        proposta = self._obter_proposta_selecionada()
        if not proposta:
            messagebox.showinfo("Informação", "Selecione uma proposta.")
            return

        win = tk.Toplevel(self)
        win.title(f"Alterar Status - Proposta #{proposta.id}")
        win.geometry("300x180")

        ttk.Label(win, text=f"Status atual: {proposta.status}")\
            .pack(anchor="w", padx=10, pady=5)
        ttk.Label(win, text="Novo status:").pack(anchor="w", padx=10, pady=5)

        status_var = tk.StringVar(value=proposta.status)
        combo_status = ttk.Combobox(
            win,
            values=Proposta.STATUS_VALIDOS,
            textvariable=status_var,
            state="readonly"
        )
        combo_status.pack(fill=tk.X, padx=10)

        def aplicar():
            novo_status = combo_status.get()
            try:
                proposta.alterar_status(novo_status)
                self.atualizar_listas()
                self.status_var.set(
                    f"Status da Proposta #{proposta.id} alterado para '{novo_status}'."
                )
                win.destroy()
            except ValueError as e:
                messagebox.showerror("Erro", str(e))

        ttk.Button(win, text="Alterar", command=aplicar).pack(pady=10)

    def acao_duplicar_proposta(self):
        proposta = self._obter_proposta_selecionada()
        if not proposta:
            messagebox.showinfo("Informação", "Selecione uma proposta.")
            return

        nova = self.gestor.criar_proposta(
            proposta.cliente,
            titulo=f"{proposta.titulo} (cópia)",
            validade=proposta.validade,
            responsavel=proposta.responsavel,
            condicoes_pagamento=proposta.condicoes_pagamento,
        )
        nova.tipo_desconto = proposta.tipo_desconto
        nova.desconto_percentual = proposta.desconto_percentual
        nova.desconto_valor = proposta.desconto_valor

        for item in proposta.itens:
            novo_item = ItemProposta(item.descricao, item.quantidade, item.valor_unitario)
            nova.adicionar_item(novo_item)

        self.atualizar_listas()
        self.status_var.set(f"Proposta #{proposta.id} duplicada como #{nova.id}.")


    def janela_adicionar_itens(self):
        proposta = self._obter_proposta_selecionada()
        if not proposta:
            messagebox.showinfo("Informação", "Selecione uma proposta.")
            return

        win = tk.Toplevel(self)
        win.title(f"Adicionar Itens - Proposta #{proposta.id}")
        win.geometry("450x260")

        ttk.Label(win, text=f"Proposta #{proposta.id} - {proposta.titulo}")\
            .pack(anchor="w", padx=10, pady=5)

        ttk.Label(win, text="Descrição do item:").pack(anchor="w", padx=10, pady=5)
        entry_desc = ttk.Entry(win)
        entry_desc.pack(fill=tk.X, padx=10)

        ttk.Label(win, text="Quantidade:").pack(anchor="w", padx=10, pady=5)
        entry_qtd = ttk.Entry(win)
        entry_qtd.pack(fill=tk.X, padx=10)

        ttk.Label(win, text="Valor unitário (R$):").pack(anchor="w", padx=10, pady=5)
        entry_valor = ttk.Entry(win)
        entry_valor.pack(fill=tk.X, padx=10)

        def adicionar():
            desc = entry_desc.get().strip()
            if not desc:
                messagebox.showwarning("Atenção", "Descrição é obrigatória.")
                return
            try:
                qtd = int(entry_qtd.get().strip())
                valor = float(entry_valor.get().strip().replace(",", "."))
            except ValueError:
                messagebox.showwarning("Atenção", "Quantidade e valor devem ser numéricos.")
                return

            item = ItemProposta(desc, qtd, valor)
            proposta.adicionar_item(item)
            self.atualizar_listas()
            self.status_var.set(f"Item adicionado à Proposta #{proposta.id}.")
            entry_desc.delete(0, tk.END)
            entry_qtd.delete(0, tk.END)
            entry_valor.delete(0, tk.END)

        ttk.Button(win, text="Adicionar Item", command=adicionar).pack(pady=10)

    def janela_editar_item(self):
        proposta = self._obter_proposta_selecionada()
        if not proposta:
            messagebox.showinfo("Informação", "Selecione uma proposta.")
            return

        if not proposta.itens:
            messagebox.showinfo("Informação", "Essa proposta ainda não possui itens.")
            return

        selecao_tree = self.tree_itens.selection()
        if not selecao_tree:
            messagebox.showinfo("Informação", "Selecione um item na lista para editar.")
            return

        item_id = selecao_tree[0]
        indice_item = self.tree_itens.index(item_id)

        if not (0 <= indice_item < len(proposta.itens)):
            messagebox.showerror("Erro", "Não foi possível localizar o item selecionado.")
            return

        item = proposta.itens[indice_item]

        win = tk.Toplevel(self)
        win.title(f"Editar Item #{indice_item + 1} - Proposta #{proposta.id}")
        win.geometry("450x260")

        ttk.Label(win, text=f"Proposta #{proposta.id} - {proposta.titulo}")\
            .pack(anchor="w", padx=10, pady=5)

        ttk.Label(win, text="Descrição do item:").pack(anchor="w", padx=10, pady=5)
        entry_desc = ttk.Entry(win)
        entry_desc.pack(fill=tk.X, padx=10)
        entry_desc.insert(0, item.descricao)

        ttk.Label(win, text="Quantidade:").pack(anchor="w", padx=10, pady=5)
        entry_qtd = ttk.Entry(win)
        entry_qtd.pack(fill=tk.X, padx=10)
        entry_qtd.insert(0, str(item.quantidade))

        ttk.Label(win, text="Valor unitário (R$):").pack(anchor="w", padx=10, pady=5)
        entry_valor = ttk.Entry(win)
        entry_valor.pack(fill=tk.X, padx=10)
        entry_valor.insert(0, f"{item.valor_unitario:.2f}".replace(".", ","))

        def salvar():
            desc = entry_desc.get().strip()
            if not desc:
                messagebox.showwarning("Atenção", "Descrição é obrigatória.")
                return
            try:
                qtd = int(entry_qtd.get().strip())
                valor = float(entry_valor.get().strip().replace(",", "."))
            except ValueError:
                messagebox.showwarning("Atenção", "Quantidade e valor devem ser numéricos.")
                return

            item.descricao = desc
            item.quantidade = qtd
            item.valor_unitario = valor

            self.atualizar_listas()
            self.status_var.set(
                f"Item #{indice_item + 1} da Proposta #{proposta.id} atualizado com sucesso."
            )
            win.destroy()

        def remover():
            if messagebox.askyesno(
                "Confirmar remoção",
                f"Remover o item #{indice_item + 1} desta proposta?"
            ):
                del proposta.itens[indice_item]
                self.atualizar_listas()
                self.status_var.set(
                    f"Item #{indice_item + 1} removido da Proposta #{proposta.id}."
                )
                win.destroy()

        frame_botoes = ttk.Frame(win)
        frame_botoes.pack(pady=10)

        ttk.Button(frame_botoes, text="Salvar alterações", command=salvar)\
            .grid(row=0, column=0, padx=5)
        ttk.Button(frame_botoes, text="Remover Item", command=remover)\
            .grid(row=0, column=1, padx=5)

    def acao_gerar_excel(self):
        if not self.gestor.listar_propostas():
            messagebox.showinfo("Informação", "Nenhuma proposta cadastrada para gerar Excel.")
            return

        default_name = f"propostas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        caminho = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            initialfile=default_name,
            title="Salvar relatório Excel"
        )
        if not caminho:
            return

        try:
            caminho_final = ExcelReportGenerator.gerar_excel(self.gestor, caminho)
            messagebox.showinfo("Sucesso", f"Arquivo Excel gerado em:\n{caminho_final}")
            self.status_var.set(f"Excel gerado em: {caminho_final}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao gerar o Excel:\n{e}")
