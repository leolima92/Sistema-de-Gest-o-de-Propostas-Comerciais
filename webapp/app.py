import os
import sys
from datetime import datetime
from functools import wraps

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    send_file,
    flash,
    session,
)

# --- Garantir que a raiz do projeto está no sys.path ---
BASE_DIR = os.path.dirname(os.path.dirname(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from gestor_propostas.models import GestorPropostas, ItemProposta
from gestor_propostas.storage import StorageManager
from gestor_propostas.excel_report import ExcelReportGenerator
from gestor_propostas.pdf_report import PdfReportGenerator
from gestor_propostas.auth import AuthManager


app = Flask(__name__)
app.secret_key = "troque-este-segredo-depois"  # precisa para session/flash

gestor = GestorPropostas()
StorageManager.init_db()
StorageManager.carregar_tudo(gestor)


def login_required(view_func):
    @wraps(view_func)
    def wrapper(*args, **kwargs):
        if "username" not in session:
            return redirect(url_for("login", next=request.path))
        return view_func(*args, **kwargs)

    return wrapper


@app.context_processor
def inject_user():
    """Disponibiliza o usuário logado no template como 'usuario_logado'."""
    return {"usuario_logado": session.get("username")}


@app.route("/login", methods=["GET", "POST"])
def login():
    
    if "username" in session:
        return redirect(url_for("index"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        user = AuthManager.authenticate(username, password)
        if user:
            session["username"] = user.username
            flash(f"Bem-vindo, {user.username}!", "success")

            next_page = request.args.get("next")
            return redirect(next_page or url_for("index"))
        else:
            flash("Usuário ou senha inválidos.", "error")

    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    session.clear()
    flash("Sessão encerrada.", "info")
    return redirect(url_for("login"))


@app.route("/register", methods=["GET", "POST"])
def register():
    if "username" in session:
        return redirect(url_for("index"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        confirm = request.form.get("confirm", "").strip()

        if not username or not password:
            flash("Usuário e senha são obrigatórios.", "error")
            return redirect(url_for("register"))

        if password != confirm:
            flash("As senhas não conferem.", "error")
            return redirect(url_for("register"))

        user = AuthManager.create_user(username, password)
        if not user:
            flash("Usuário já existe. Escolha outro.", "error")
            return redirect(url_for("register"))

        flash("Usuário criado com sucesso! Faça login.", "success")
        return redirect(url_for("login"))

    return render_template("register.html")


@app.route("/")
@login_required
def index():
    q = request.args.get("q", "").strip().lower()
    status = request.args.get("status", "").strip()

    todas_propostas = gestor.listar_propostas()
    propostas = todas_propostas

    if status:
        propostas = [p for p in propostas if p.status == status]

    if q:
        propostas = [
            p for p in propostas
            if q in p.titulo.lower() or q in p.cliente.nome.lower()
        ]


    statuses = sorted({p.status for p in todas_propostas})

    total_propostas = len(todas_propostas)
    total_clientes = len(gestor.listar_clientes())
    propostas_aceitas = [p for p in todas_propostas if p.status == "aceita"]
    qtd_aceitas = len(propostas_aceitas)
    valor_total_aceitas = sum(p.calcular_total() for p in propostas_aceitas)

    return render_template(
        "index.html",
        propostas=propostas,
        filtro_q=q,
        filtro_status=status,
        statuses=statuses,
        total_propostas=total_propostas,
        total_clientes=total_clientes,
        qtd_aceitas=qtd_aceitas,
        valor_total_aceitas=valor_total_aceitas,
    )




@app.route("/propostas/<int:pid>")
@login_required
def proposta_detalhe(pid: int):
    """Detalhe de uma proposta específica."""
    proposta = next((p for p in gestor.propostas if p.id == pid), None)
    if not proposta:
        flash("Proposta não encontrada.", "error")
        return redirect(url_for("index"))

    return render_template("proposta_detalhe.html", proposta=proposta)


@app.route("/propostas/nova", methods=["GET", "POST"])
@login_required
def nova_proposta():
    """Criação de uma nova proposta."""
    clientes = gestor.listar_clientes()

    if not clientes:
        flash("Cadastre ao menos um cliente antes de criar uma proposta.", "info")
        return redirect(url_for("novo_cliente"))

    if request.method == "POST":
        cliente_id = int(request.form.get("cliente_id", "0"))
        titulo = request.form.get("titulo", "").strip()
        responsavel = request.form.get("responsavel", "").strip()
        validade_str = request.form.get("validade", "").strip()
        cond_pag = request.form.get("condicoes_pagamento", "").strip()

        cliente = next((c for c in clientes if c.id == cliente_id), None)
        if not cliente:
            flash("Cliente inválido.", "error")
            return redirect(url_for("nova_proposta"))

        validade = None
        if validade_str:
            try:
                validade = datetime.strptime(validade_str, "%Y-%m-%d").date()
            except ValueError:
                flash("Data de validade inválida. Use AAAA-MM-DD.", "error")
                return redirect(url_for("nova_proposta"))

        prop = gestor.criar_proposta(
            cliente,
            titulo,
            validade=validade,
            responsavel=responsavel,
            condicoes_pagamento=cond_pag,
        )

        StorageManager.salvar_ou_atualizar_proposta(prop)
        StorageManager.sincronizar_itens_proposta(prop)

        flash(f"Proposta #{prop.id} criada com sucesso!", "success")
        return redirect(url_for("proposta_detalhe", pid=prop.id))

    # GET
    return render_template("nova_proposta.html", clientes=clientes)


@app.route("/propostas/<int:pid>/add_item", methods=["POST"])
@login_required
def add_item(pid: int):
    """Adicionar item a uma proposta."""
    proposta = next((p for p in gestor.propostas if p.id == pid), None)
    if not proposta:
        flash("Proposta não encontrada.", "error")
        return redirect(url_for("index"))

    desc = request.form.get("descricao", "").strip()
    qtd_str = request.form.get("quantidade", "0").strip()
    valor_str = request.form.get("valor_unitario", "0").strip()

    if not desc:
        flash("Descrição é obrigatória.", "error")
        return redirect(url_for("proposta_detalhe", pid=pid))

    try:
        qtd = int(qtd_str)
        valor = float(valor_str.replace(",", "."))
    except ValueError:
        flash("Quantidade e valor devem ser numéricos.", "error")
        return redirect(url_for("proposta_detalhe", pid=pid))

    item = ItemProposta(desc, qtd, valor)
    proposta.adicionar_item(item)

    StorageManager.sincronizar_itens_proposta(proposta)
    StorageManager.salvar_ou_atualizar_proposta(proposta)

    flash("Item adicionado com sucesso!", "success")
    return redirect(url_for("proposta_detalhe", pid=pid))


@app.route("/propostas/<int:pid>/desconto", methods=["POST"])
@login_required
def aplicar_desconto(pid: int):
   
    proposta = next((p for p in gestor.propostas if p.id == pid), None)
    if not proposta:
        flash("Proposta não encontrada.", "error")
        return redirect(url_for("index"))

    tipo = request.form.get("tipo", "nenhum")
    valor_str = request.form.get("valor", "").strip()

    if tipo == "nenhum":
        proposta.tipo_desconto = None
        proposta.desconto_percentual = 0.0
        proposta.desconto_valor = 0.0
        msg = "Desconto removido."
    else:
        try:
            valor = float(valor_str.replace(",", "."))
        except ValueError:
            flash("Informe um valor numérico para desconto.", "error")
            return redirect(url_for("proposta_detalhe", pid=pid))

        if tipo == "%":
            proposta.definir_desconto_percentual(valor)
            msg = f"Desconto de {valor:.2f}% aplicado."
        elif tipo == "R":
            proposta.definir_desconto_valor(valor)
            msg = f"Desconto de R$ {valor:.2f} aplicado."
        else:
            msg = "Tipo de desconto inválido."

    StorageManager.salvar_ou_atualizar_proposta(proposta)
    flash(msg, "success")
    return redirect(url_for("proposta_detalhe", pid=pid))


@app.route("/propostas/<int:pid>/excluir", methods=["POST"])
@login_required
def excluir_proposta(pid: int):
    proposta = next((p for p in gestor.propostas if p.id == pid), None)
    if not proposta:
        flash("Proposta não encontrada.", "error")
        return redirect(url_for("index"))

    gestor.propostas = [p for p in gestor.propostas if p.id != pid]

    try:
        if hasattr(StorageManager, "excluir_proposta"):
            StorageManager.excluir_proposta(pid)
        elif hasattr(StorageManager, "deletar_proposta"):
            StorageManager.deletar_proposta(pid)
    except Exception:
        
        flash("Erro ao excluir no banco, mas proposta foi removida da lista atual.", "error")

    flash(f"Proposta #{pid} excluída com sucesso.", "success")
    return redirect(url_for("index"))


@app.route("/propostas/<int:pid>/enviar", methods=["POST"])
@login_required
def enviar_proposta(pid: int):
    proposta = next((p for p in gestor.propostas if p.id == pid), None)
    if not proposta:
        flash("Proposta não encontrada.", "error")
        return redirect(url_for("index"))

    if proposta.status in ["enviada", "aceita", "recusada", "cancelada"]:
        flash(f"A Proposta #{pid} já está com status '{proposta.status}'.", "info")
        return redirect(url_for("index"))

    proposta.alterar_status("enviada")
    StorageManager.salvar_ou_atualizar_proposta(proposta)

    flash(f"Proposta #{pid} marcada como 'enviada'.", "success")
    return redirect(url_for("index"))



@app.route("/propostas/<int:pid>/excel")
@login_required
def download_excel(pid: int):
    """Gera um Excel com todas as propostas e manda para download."""
    if not gestor.listar_propostas():
        flash("Não há propostas para exportar.", "info")
        return redirect(url_for("index"))

    caminho = ExcelReportGenerator.gerar_excel(gestor)
    return send_file(caminho, as_attachment=True)


@app.route("/propostas/<int:pid>/pdf")
@login_required
def download_pdf(pid: int):
    proposta = next((p for p in gestor.propostas if p.id == pid), None)
    if not proposta:
        flash("Proposta não encontrada.", "error")
        return redirect(url_for("index"))

    from tempfile import NamedTemporaryFile

    tmp = NamedTemporaryFile(delete=False, suffix=f"_proposta_{proposta.id}.pdf")
    tmp.close()

    PdfReportGenerator.gerar_pdf_proposta(proposta, tmp.name)

    return send_file(tmp.name, as_attachment=True, download_name=f"proposta_{proposta.id}.pdf")

@app.route("/clientes")
@login_required
def clientes():
    return render_template("clientes.html", clientes=gestor.listar_clientes())


@app.route("/clientes/novo", methods=["GET", "POST"])
@login_required
def novo_cliente():
    if request.method == "POST":
        nome = request.form.get("nome", "").strip()
        documento = request.form.get("documento", "").strip()
        contato = request.form.get("contato", "").strip()

        if not nome:
            flash("Nome do cliente é obrigatório.", "error")
            return redirect(url_for("novo_cliente"))

        cliente = gestor.criar_cliente(nome, documento, contato)
        StorageManager.salvar_ou_atualizar_cliente(cliente)

        flash(f"Cliente '{cliente.nome}' criado com sucesso!", "success")
        return redirect(url_for("clientes"))

    return render_template("novo_cliente.html")


if __name__ == "__main__":
    app.run(debug=True)
