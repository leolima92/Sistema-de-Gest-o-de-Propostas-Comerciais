document.addEventListener("DOMContentLoaded", () => {
    const html = document.documentElement;
    const btn = document.getElementById("themeToggle");

    // Carrega tema salvo
    let saved = localStorage.getItem("theme");
    if (saved) {
        html.setAttribute("data-theme", saved);
    }

    // Atualiza ícone inicial
    updateIcon();

    btn.addEventListener("click", () => {
        let current = html.getAttribute("data-theme");

        let newTheme = current === "dark" ? "light" : "dark";
        html.setAttribute("data-theme", newTheme);

        localStorage.setItem("theme", newTheme);
        updateIcon();
    });

    function updateIcon() {
        let current = html.getAttribute("data-theme");
        btn.textContent = current === "dark" ? "☀️" : "🌙";
    }
    
    document.querySelectorAll('a[href^="#"]').forEach(link => {
        link.addEventListener("click", e => {
            const targetId = link.getAttribute("href").substring(1);
            const target = document.getElementById(targetId);
            if (target) {
                e.preventDefault();
                target.scrollIntoView({ behavior: "smooth" });
            }
        });
    });
});
