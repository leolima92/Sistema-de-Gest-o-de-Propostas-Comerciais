from gestor_propostas.models import GestorPropostas
from gestor_propostas.ui import App
from gestor_propostas.login_ui import LoginWindow
from gestor_propostas.storage import StorageManager


def main():
    login = LoginWindow()
    login.mainloop()

    user = login.user
    if not user:
        return

    gestor = GestorPropostas()
    StorageManager.init_db()
    StorageManager.carregar_tudo(gestor)

    app = App(gestor, usuario_logado=user.username)
    app.atualizar_listas()

    def on_close():
        try:
            StorageManager.salvar_tudo(gestor)
        finally:
            app.destroy()

    app.protocol("WM_DELETE_WINDOW", on_close)
    app.mainloop()


if __name__ == "__main__":
    main()
