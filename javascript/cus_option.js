$(document).ready(function () {
    const newlinkBtn2 = document.querySelector('.burger');
    const newlinkList2 = document.querySelector('.nav__menu');
    const childrenList2 = document.querySelectorAll('.nav__list>li.nav__item')
    newlinkBtn2.addEventListener('click', () => {
        const iClass = document.querySelector('div.burger i');
        iClass.classList.toggle('fa-times');
        iClass.classList.toggle('fa-bars');
        console.log(iClass);
        newlinkList2.classList.toggle('show-burger');
        childrenList2.forEach((element, index) => {
            element.style.transition = `all 0.7s ease ${index / 30}s`;
            element.classList.toggle('transitionForLi');
        })
    });
    // click to show under the chevron1 
    const navMoblie = document.querySelector('ul.nav__list');
    navMoblie.addEventListener('click', (e) => {
        e.preventDefault;
        const currentTarget = e.target.closest('.dropdown__link');
        const showUlMoblie = currentTarget.nextElementSibling;
        const childrenFlop = Array.from(showUlMoblie.children);
        const arrow = currentTarget.firstElementChild;
        if (!currentTarget) return;
        console.log(childrenFlop);
        childrenFlop.forEach((element, index) => {
            element.style.transition = `all 0.7s ease ${index / 30}s`;
            element.classList.toggle('transitionForLi');
        });
        showUlMoblie.classList.toggle('height');
        arrow.classList.toggle('bx-chevron-down-reverse')

    })
    var owl = $("#home .owl-carousel");
    owl.owlCarousel({
        items: 1,
        loop: true,
        nav: false,
        margin: 10,
        autoplay: true,
        autoplayTimeout: 2500,
        autoplayHoverPause: true,
        lazyLoad: true,
        dots: true,
    });
    window.addEventListener("load", () => {
        const owlCarousel = document.getElementById("home");
        owlCarousel.scrollIntoView({ behavior: "smooth", block: "end", inline: "center" });
    });
    const closeSupport = document.querySelector(".supportContainer .top");
    const supportContainer = document.querySelector(".supportContainer");
    closeSupport.addEventListener("click", () => {
        supportContainer.classList.add("hide");
    });
    const openSupport = document.querySelector("#style-switcher-toggle");
    openSupport.addEventListener("click", () => {
        supportContainer.classList.remove("hide");
    });
});
