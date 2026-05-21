(function () {
  function getEffectName(element) {
    return element.getAttribute("swiper-animate-effect") || element.dataset.effect || "fadeInUp";
  }

  function getDuration(element) {
    return element.getAttribute("swiper-animate-duration") || "0.7s";
  }

  function getDelay(element) {
    return element.getAttribute("swiper-animate-delay") || "0s";
  }

  function resetAnimations(container) {
    container.querySelectorAll(".ani").forEach(function (element) {
      var effect = getEffectName(element);
      element.classList.remove("animated", "effect-" + effect);
      element.style.animationDuration = "";
      element.style.animationDelay = "";
      element.style.animationFillMode = "";
    });
  }

  function runAnimations(slide) {
    slide.querySelectorAll(".ani").forEach(function (element) {
      var effect = getEffectName(element);
      element.classList.remove("animated", "effect-" + effect);
      void element.offsetWidth;
      element.style.animationDuration = getDuration(element);
      element.style.animationDelay = getDelay(element);
      element.style.animationFillMode = "both";
      element.classList.add("animated", "effect-" + effect);
    });
  }

  function fallbackSwiper() {
    var root = document.querySelector(".invite-swiper");
    var wrapper = root.querySelector(".swiper-wrapper");
    var slides = Array.from(root.querySelectorAll(".swiper-slide"));
    var index = Number(new URLSearchParams(window.location.search).get("slide") || 0);
    index = Math.max(0, Math.min(slides.length - 1, index));

    function show(nextIndex) {
      index = (nextIndex + slides.length) % slides.length;
      wrapper.style.transform = "translate3d(0, -" + index * (100 / slides.length) + "%, 0)";
      wrapper.style.transition = "transform 420ms ease";
      resetAnimations(root);
      runAnimations(slides[index]);
    }

    root.classList.add("fallback-swiper");
    wrapper.style.height = slides.length * 100 + "%";
    slides.forEach(function (slide) {
      slide.style.height = 100 / slides.length + "%";
    });
    root.querySelector(".slide-tip").addEventListener("click", function () {
      show(index + 1);
    });
    root.addEventListener("wheel", function (event) {
      event.preventDefault();
      show(index + (event.deltaY > 0 ? 1 : -1));
    }, { passive: false });
    show(index);
  }

  function initSwiper() {
    var initialSlide = Number(new URLSearchParams(window.location.search).get("slide") || 0);

    if (!window.Swiper) {
      fallbackSwiper();
      return;
    }

    var swiper = new Swiper(".invite-swiper", {
      direction: "vertical",
      loop: false,
      initialSlide: initialSlide,
      speed: 620,
      mousewheel: true,
      pagination: {
        el: ".swiper-pagination",
        clickable: true
      },
      on: {
        init: function () {
          resetAnimations(document);
          runAnimations(this.slides[this.activeIndex]);
        },
        slideChangeTransitionStart: function () {
          resetAnimations(document);
        },
        slideChangeTransitionEnd: function () {
          runAnimations(this.slides[this.activeIndex]);
        }
      }
    });

    document.querySelector(".slide-tip").addEventListener("click", function () {
      swiper.slideNext();
    });
  }

  window.addEventListener("load", initSwiper);
})();
