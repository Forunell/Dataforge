document.querySelectorAll('input[type="file"]').forEach(input => {
  const zone = input.closest('.dropzone');
  const update = () => {
    const file = input.files[0];
    input.setCustomValidity(file && file.size > 2 * 1024 * 1024 ? 'Максимальний розмір файлу — 2 МБ.' : '');
    zone.querySelector('[data-filename]').textContent = file ? `${file.name} · ${file.size.toLocaleString('uk-UA')} байтів` : 'Виберіть файл';
    zone.classList.toggle('selected', !!file);
  };
  input.addEventListener('change', update);
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragging'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('dragging'));
  zone.addEventListener('drop', e => {
    e.preventDefault(); zone.classList.remove('dragging');
    if (e.dataTransfer.files.length === 1) { input.files = e.dataTransfer.files; update(); }
  });
});
document.querySelectorAll('form').forEach(form => form.addEventListener('submit', () => {
  document.getElementById('busy').hidden = false;
  form.querySelectorAll('button[type="submit"]').forEach(button => { button.disabled = true; });
}));
window.addEventListener('pageshow', () => {
  document.getElementById('busy').hidden = true;
  document.querySelectorAll('button[type="submit"]').forEach(button => { button.disabled = false; });
});
document.querySelectorAll('select[name$="_sku"], select[name$="_name"], select[name$="_price"]').forEach(select => {
  const validate = () => select.setCustomValidity(select.value === '-1' ? 'Виберіть колонку.' : '');
  select.addEventListener('change', validate);
  validate();
});
if (location.hash === '#help') document.getElementById('help')?.setAttribute('open', '');
window.addEventListener('hashchange', () => {
  if (location.hash === '#help') document.getElementById('help')?.setAttribute('open', '');
});
