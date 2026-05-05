(() => {
  document.addEventListener("DOMContentLoaded", initToolbox);

  function initToolbox() {
    const tools = Object.values(window.FAFULITools || {});
    const nav = document.getElementById("toolNav");
    const panels = document.getElementById("toolPanels");
    if (!tools.length || !nav || !panels) return;

    renderNav(nav, tools);
    renderPanels(panels, tools);
    tools.forEach((tool) => {
      const panel = panels.querySelector(`[data-tool-panel="${tool.id}"]`);
      tool.init(panel);
    });
    activateTool(tools[0].id);
  }

  function renderNav(nav, tools) {
    const groups = [];
    tools.forEach((tool) => {
      let group = groups.find((item) => item.name === tool.group);
      if (!group) {
        group = { name: tool.group, tools: [] };
        groups.push(group);
      }
      group.tools.push(tool);
    });

    nav.innerHTML = groups.map((group) => `
      <p class="nav-group-title">${escapeHtml(group.name)}</p>
      ${group.tools.map((tool) => `
        <div class="nav-item" data-tool-nav="${escapeHtml(tool.id)}">${escapeHtml(tool.name)}</div>
      `).join("")}
    `).join("");

    nav.querySelectorAll("[data-tool-nav]").forEach((item) => {
      item.addEventListener("click", () => activateTool(item.dataset.toolNav));
    });
  }

  function renderPanels(panels, tools) {
    panels.innerHTML = tools.map((tool) => `
      <section class="tool-panel" data-tool-panel="${escapeHtml(tool.id)}">${tool.render()}</section>
    `).join("");
  }

  function activateTool(toolId) {
    document.querySelectorAll("[data-tool-nav]").forEach((navItem) => {
      const active = navItem.dataset.toolNav === toolId;
      navItem.classList.toggle("nav-item-active", active);
      if (active) {
        navItem.setAttribute("aria-current", "page");
      } else {
        navItem.removeAttribute("aria-current");
      }
    });

    document.querySelectorAll("[data-tool-panel]").forEach((panel) => {
      panel.classList.toggle("is-active", panel.dataset.toolPanel === toolId);
    });
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }
})();
