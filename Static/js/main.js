document.querySelectorAll("input").forEach(el => {
  el.addEventListener("focus", () => {
    el.style.background = "#fff8f9";
  });
  el.addEventListener("blur", () => {
    el.style.background = "#fff";
  });
});
