/* Sales Intelligence sidebar / navigation behavior */
(function () {
    const body = document.body;
    const sidebar = document.getElementById('siSidebar');
    const toggle = document.getElementById('sidebarToggle');
    const mobileToggle = document.getElementById('mobileMenuToggle');
    const backdrop = document.getElementById('sidebarBackdrop');

    if (!sidebar) return;

    const isMobile = () => window.matchMedia('(max-width: 991.98px)').matches;

    function setCollapsed(collapsed) {
        if (isMobile()) return;

        body.classList.toggle('sidebar-collapsed', collapsed);

        if (toggle) {
            toggle.setAttribute('aria-expanded', String(!collapsed));
            toggle.setAttribute('aria-label', collapsed ? 'Expand sidebar' : 'Collapse sidebar');
            toggle.title = collapsed ? 'Expand sidebar' : 'Collapse sidebar';
            toggle.innerHTML = collapsed
                ? '<i class="bi bi-layout-sidebar"></i>'
                : '<i class="bi bi-layout-sidebar-inset"></i>';
        }

        localStorage.setItem('si-sidebar-collapsed', collapsed ? '1' : '0');
    }

    function openMobile() {
        body.classList.add('sidebar-mobile-open');
        mobileToggle?.setAttribute('aria-expanded', 'true');
    }

    function closeMobile() {
        body.classList.remove('sidebar-mobile-open');
        mobileToggle?.setAttribute('aria-expanded', 'false');
    }

    // Restore desktop collapsed state.
    const saved = localStorage.getItem('si-sidebar-collapsed') === '1';
    if (!isMobile()) setCollapsed(saved);

    toggle?.addEventListener('click', () => {
        setCollapsed(!body.classList.contains('sidebar-collapsed'));
    });

    mobileToggle?.addEventListener('click', () => {
        body.classList.contains('sidebar-mobile-open') ? closeMobile() : openMobile();
    });

    backdrop?.addEventListener('click', closeMobile);

    document.querySelectorAll('.si-nav-link').forEach(link => {
        link.addEventListener('click', closeMobile);
    });

    // Highlight current route.
    const path = window.location.pathname.replace(/\/$/, '') || '/';
    document.querySelectorAll('.si-nav-link').forEach(link => {
        const route = link.dataset.route || '';
        const normalized = route.replace(/\/$/, '') || '/';
        link.classList.toggle('active', path === normalized);
    });

    window.addEventListener('resize', () => {
        if (isMobile()) {
            body.classList.remove('sidebar-collapsed');
        } else {
            closeMobile();
            setCollapsed(localStorage.getItem('si-sidebar-collapsed') === '1');
        }
    });

    // User initials are shown in both sidebar and top navigation.
    const userName = document.getElementById('siUserName')?.textContent.trim() || 'User';
    const userEmail = document.getElementById('siUserEmail')?.textContent.trim() || '';
    const displayName = userName !== 'Logged-in User' ? userName : (userEmail || userName);

    const initials = displayName
        .split(/\s+/)
        .filter(Boolean)
        .slice(0, 2)
        .map(part => part[0].toUpperCase())
        .join('') || 'U';

    ['siUserAvatar', 'siTopAvatar'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.textContent = initials;
    });

    const topName = document.getElementById('siTopUserName');
    if (topName && userName) topName.textContent = userName;
})();
