
/**
 * 設置滾動監聽 (Scroll Spy)
 */
function setupScrollSpy() {
    console.log('Setup Scroll Spy...');
    const mainContent = document.querySelector('main');
    // Using IntersectionObserver with a larger threshold for better accuracy
    const observerOptions = {
        root: mainContent,
        threshold: 0.5 // 50% visibility
    };

    const sections = ['step1', 'step2', 'step3'];
    const navLinks = {
        'step1': document.getElementById('nav-step1'),
        'step2': document.getElementById('nav-step2'),
        'step3': document.getElementById('nav-step3')
    };

    // Check if nav links exist before proceeding
    if (!navLinks['step1'] || !navLinks['step2'] || !navLinks['step3']) {
        console.warn('Scroll Spy: Navigation links not found in DOM.');
        return;
    }

    const setActive = (activeId) => {
        // User requested to disable active state highlighting on scroll.
        // Only hover effects should change the appearance.
        return;
    };

    const observer = new IntersectionObserver((entries) => {
        // Observer kept running but doing nothing for now, 
        // in case we want to re-enable logic later or use it for other things.
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                setActive(entry.target.id);
            }
        });
    }, observerOptions);

    sections.forEach(id => {
        const section = document.getElementById(id);
        if (section) observer.observe(section);
    });

    // Smooth Scrolling for Nav Links
    Object.values(navLinks).forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const targetId = link.getAttribute('href').substring(1);
            const targetSection = document.getElementById(targetId);
            if (targetSection) {
                targetSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        });
    });
}
