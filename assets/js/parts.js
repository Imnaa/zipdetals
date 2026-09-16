/*
 *  Таблица запчастей: данные берутся из Google Таблицы (или из резервного
 *  CSV в репозитории), а не из разметки. Правьте таблицу — см. README.md.
 */
(function () {
	'use strict';

	/* ------------------------------------------------------------------
	 *  НАСТРОЙКА
	 *
	 *  Основной источник — Google Таблица (правится прямо в браузере,
	 *  сайт подхватывает изменения сразу, без коммита в GitHub):
	 *  https://docs.google.com/spreadsheets/d/1kIhsasNJc_ByQAFeM96MynhXWAjgqAsnrTY4PU_g9JM/edit
	 *
	 *  Таблица должна быть открыта на чтение: Доступ → «Все, у кого есть
	 *  ссылка» → Читатель. Чтобы переехать на другую таблицу, достаточно
	 *  заменить SHEET_ID — это часть ссылки между /d/ и /edit.
	 *
	 *  Если Google недоступен, список берётся из data/parts.csv в репозитории.
	 * ------------------------------------------------------------------ */
	var SHEET_ID = '1kIhsasNJc_ByQAFeM96MynhXWAjgqAsnrTY4PU_g9JM';
	var SHEET_URL = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/gviz/tq?tqx=out:csv';
	var BACKUP_URL = 'data/parts.csv';

	// Куда уходит письмо с заказом.
	var MAIL_TO = 'pershin1950@mail.ru';
	var MAIL_CC = 'kolya.romashenko@yandex.ru';
	var SITE_URL = 'https://imnaa.github.io/zipdetals/';

	var STORAGE_KEY = 'zipdetals.order.v1';   // выбранные позиции переживают перезагрузку
	var MAILTO_LIMIT = 2000;                  // дальше Outlook и часть клиентов режут ссылку

	// Первая колонка — выбор; заголовок короткий, иначе колонка раздувается.
	var COLUMNS = ['\u2713', '№', 'Наименование', 'Код/чертеж', 'Цена за шт.', 'Кол-во, шт.', 'Сумма', 'Фото'];
	var PHOTO_SEPARATOR = /\s*[|;]\s*/;

	/* ---------------------------- CSV ---------------------------------- */

	// Разбор CSV с поддержкой кавычек и переносов строк внутри ячеек.
	function parseCSV(text, delimiter) {
		var rows = [], row = [], field = '', inQuotes = false, i = 0;

		if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);   // BOM из Excel

		while (i < text.length) {
			var ch = text[i];

			if (inQuotes) {
				if (ch === '"') {
					if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
					inQuotes = false; i++; continue;
				}
				field += ch; i++; continue;
			}

			if (ch === '"') { inQuotes = true; i++; continue; }
			if (ch === delimiter) { row.push(field); field = ''; i++; continue; }
			if (ch === '\r') { i++; continue; }
			if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }

			field += ch; i++;
		}
		row.push(field);
		rows.push(row);

		return rows.filter(function (r) {
			return r.some(function (c) { return c.trim() !== ''; });
		});
	}

	// Excel в русской локали сохраняет с ";", Google отдаёт с ",".
	function detectDelimiter(text) {
		var head = text.split('\n')[0];
		return (head.split(';').length > head.split(',').length) ? ';' : ',';
	}

	/* --------------------- разбор колонок таблицы ----------------------- */

	// Колонки ищем по названию, а не по порядку: их можно двигать и
	// переименовывать в разумных пределах, не трогая код.
	function norm(s) {
		return String(s).toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9]/g, '');
	}

	var COLUMN_RULES = {
		category: function (k) { return k.indexOf('категор') === 0; },
		section:  function (k) { return k.indexOf('раздел') === 0 || k.indexOf('группа') === 0; },
		name:     function (k) { return k.indexOf('наимен') === 0 || k.indexOf('назван') === 0 || k === 'товар'; },
		code:     function (k) { return k.indexOf('код') === 0 || k.indexOf('чертеж') !== -1 || k.indexOf('артикул') === 0; },
		price:    function (k) { return k.indexOf('цена') === 0 || k.indexOf('стоим') === 0; },
		stock:    function (k) { return k.indexOf('колич') === 0 || k.indexOf('колво') === 0 || k.indexOf('остат') === 0; },
		photo:    function (k) { return k.indexOf('фото') === 0 || k.indexOf('изобр') === 0 || k.indexOf('картин') === 0; }
	};

	function mapColumns(header) {
		var map = {};
		header.forEach(function (title, i) {
			var k = norm(title);
			Object.keys(COLUMN_RULES).forEach(function (field) {
				if (map[field] === undefined && k && COLUMN_RULES[field](k)) map[field] = i;
			});
		});
		return map;
	}

	// "1 500,50 руб." -> 1500.5; пустое или нечисловое -> null.
	function parseNumber(v) {
		if (v == null) return null;
		var s = String(v).replace(/[\s\u00A0]/g, '').replace(/[^\d.,-]/g, '');
		if (!s) return null;

		if (s.indexOf(',') !== -1 && s.indexOf('.') !== -1) {
			// последний разделитель считаем десятичным
			s = (s.lastIndexOf(',') > s.lastIndexOf('.'))
				? s.replace(/\./g, '').replace(',', '.')
				: s.replace(/,/g, '');
		} else if (s.indexOf(',') !== -1) {
			s = s.replace(',', '.');
		}

		var n = parseFloat(s);
		return isFinite(n) ? n : null;
	}

	function toItems(rows) {
		var map = mapColumns(rows[0].map(function (h) { return h.trim(); }));
		if (map.name === undefined) throw new Error('не найдена колонка «Наименование»');

		return rows.slice(1).map(function (r) {
			function cell(field) {
				return map[field] === undefined ? '' : (r[map[field]] || '').trim();
			}
			var stock = parseNumber(cell('stock'));

			return {
				category: cell('category') || 'Прочее',
				section:  cell('section'),
				name:     cell('name'),
				code:     cell('code'),
				price:    parseNumber(cell('price')),
				stock:    stock === null ? 0 : Math.max(0, Math.floor(stock)),
				photos:   cell('photo').split(PHOTO_SEPARATOR).filter(Boolean)
			};
		}).filter(function (it) { return it.name; });
	}

	/* --------------------------- вспомогательное ------------------------ */

	function esc(s) {
		return String(s).replace(/[&<>"']/g, function (c) {
			return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
		});
	}

	function el(tag, attrs, html) {
		var n = document.createElement(tag);
		if (attrs) Object.keys(attrs).forEach(function (k) { n.setAttribute(k, attrs[k]); });
		if (html != null) n.innerHTML = html;
		return n;
	}

	// Группировка с сохранением порядка появления в файле.
	function groupBy(items, keyFn) {
		var order = [], map = {};
		items.forEach(function (it) {
			var k = keyFn(it);
			if (!(k in map)) { map[k] = []; order.push(k); }
			map[k].push(it);
		});
		return order.map(function (k) { return { name: k, items: map[k] }; });
	}

	function num(n) {
		return n.toLocaleString('ru-RU', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
	}

	function money(n) { return num(n) + ' ₽'; }        // для страницы
	function moneyPlain(n) { return num(n) + ' руб.'; }     // для письма

	/* ------------------------- выбранные позиции ------------------------ */

	var state = Object.create(null);   // ключ позиции -> { q: количество, c: выбрана }

	function keyOf(item) {
		return [item.category, item.section, item.name, item.code].join('\u0000');
	}

	function stateOf(item) {
		var k = keyOf(item);
		if (!state[k]) state[k] = { q: 1, c: false };
		return state[k];
	}

	function loadState() {
		try {
			var saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
			Object.keys(saved).forEach(function (k) {
				var v = saved[k];
				if (v && typeof v === 'object') state[k] = { q: Math.max(1, +v.q || 1), c: !!v.c };
			});
		} catch (e) {
			// приватный режим или повреждённые данные — просто начинаем с пустого выбора
		}
	}

	function saveState() {
		try {
			var out = {};
			Object.keys(state).forEach(function (k) {
				if (state[k].c) out[k] = { q: state[k].q, c: true };
			});
			localStorage.setItem(STORAGE_KEY, JSON.stringify(out));
		} catch (e) { /* хранилище недоступно — выбор просто не переживёт перезагрузку */ }
	}

	/* ------------------------------ таблица ----------------------------- */

	var allRows = [];        // { item, st, tr, checkbox, input, sumCell }
	var recalc = function () {};

	function buildRow(item, index) {
		var st = stateOf(item);
		var inStock = item.stock > 0;

		if (st.q > item.stock) st.q = Math.max(1, item.stock);
		if (!inStock) st.c = false;

		var tr = el('tr');
		tr.dataset.search = [item.name, item.code, item.section].join(' ').toLowerCase();

		// --- выбор ---
		// Галочка лежит внутри label во всю ячейку, поэтому нажатие
		// срабатывает по всему квадрату, а не только по самому квадратику.
		var pickCell = el('td', { 'class': 'col-pick' });
		var checkbox = null;
		if (inStock) {
			var pickLabel = el('label', { 'class': 'pick-box' });
			checkbox = el('input', { type: 'checkbox' });
			checkbox.checked = st.c;
			checkbox.setAttribute('aria-label', 'Выбрать: ' + item.name);
			pickLabel.appendChild(checkbox);
			pickCell.appendChild(pickLabel);
		}
		tr.appendChild(pickCell);

		tr.appendChild(el('td', { 'class': 'col-num' }, String(index)));
		tr.appendChild(el('td', null, esc(item.name)));
		tr.appendChild(el('td', null, esc(item.code)));
		tr.appendChild(el('td', { 'class': 'col-price' },
			item.price === null ? '&mdash;' : esc(money(item.price))));

		// --- количество: стрелки, не больше остатка ---
		var qtyCell = el('td', { 'class': 'col-qty' });
		var input = null;

		if (inStock) {
			var stepper = el('div', { 'class': 'qty' });
			var minus = el('button', { type: 'button', 'class': 'qty-btn', 'aria-label': 'Уменьшить количество' }, '&minus;');
			input = el('input', {
				type: 'number', 'class': 'qty-input',
				min: '1', max: String(item.stock), step: '1',
				inputmode: 'numeric', 'aria-label': 'Количество: ' + item.name
			});
			input.value = st.q;
			var plus = el('button', { type: 'button', 'class': 'qty-btn', 'aria-label': 'Увеличить количество' }, '+');

			stepper.appendChild(minus);
			stepper.appendChild(input);
			stepper.appendChild(plus);
			qtyCell.appendChild(stepper);
			qtyCell.appendChild(el('span', { 'class': 'qty-stock' }, 'из ' + item.stock));

			// Правка количества автоматически отмечает позицию.
			var apply = function (q) {
				st.q = Math.min(item.stock, Math.max(1, Math.floor(q) || 1));
				st.c = true;
				checkbox.checked = true;
				saveState();
				recalc();
			};

			minus.addEventListener('click', function () { apply(st.q - 1); input.value = st.q; });
			plus.addEventListener('click', function () { apply(st.q + 1); input.value = st.q; });

			// во время набора значение не трогаем, на blur приводим в границы
			input.addEventListener('input', function () {
				var v = parseInt(input.value, 10);
				if (isFinite(v) && v >= 1 && v <= item.stock) apply(v);
			});
			input.addEventListener('blur', function () {
				apply(parseInt(input.value, 10));
				input.value = st.q;
			});

			checkbox.addEventListener('change', function () {
				st.c = checkbox.checked;
				saveState();
				recalc();
			});
		} else {
			qtyCell.appendChild(el('span', { 'class': 'qty-out' }, 'нет в наличии'));
		}
		tr.appendChild(qtyCell);

		var sumCell = el('td', { 'class': 'col-sum' });
		tr.appendChild(sumCell);

		var photoCell = el('td', { 'class': 'col-photo' });
		if (item.photos.length) {
			photoCell.appendChild(el('button', {
				type: 'button',
				'class': 'btn btn-primary',
				'data-photos': item.photos.join('|'),
				'data-title': item.name
			}, 'Фото'));
		}
		tr.appendChild(photoCell);

		allRows.push({ item: item, st: st, tr: tr, checkbox: checkbox, input: input, sumCell: sumCell });
		return tr;
	}

	function buildTable(itemsOfCategory) {
		var table = el('table', { 'class': 'parts-table' });

		var head = el('tr');
		COLUMNS.forEach(function (c, i) {
			var attrs = { style: 'text-align: center;', 'class': 'first' };
			if (i === 0) {
				attrs['class'] += ' col-pick';
				attrs.title = 'Выбор позиции для заказа';
			}
			head.appendChild(el('th', attrs, esc(c)));
		});
		table.appendChild(head);

		groupBy(itemsOfCategory, function (it) { return it.section; }).forEach(function (section) {
			if (section.name) {
				var tr = el('tr', { 'data-section': '1' });
				tr.appendChild(el('th', {
					colspan: COLUMNS.length,
					style: 'text-align: center;',
					'class': 'first'
				}, esc(section.name) + ':'));
				table.appendChild(tr);
			}
			section.items.forEach(function (item, i) {
				table.appendChild(buildRow(item, i + 1));
			});
		});

		var wrap = el('div', { 'class': 'table-wrap' });
		wrap.appendChild(table);
		return wrap;
	}

	function buildTabs(categories) {
		var tabs = el('div', { 'class': 'tabs' });
		var panels = [];

		// Сначала все переключатели с подписями — иначе подписи неоткрытых вкладок
		// оказываются под таблицей открытой вкладки.
		categories.forEach(function (cat, i) {
			var id = 'tab-btn-' + (i + 1);
			var input = el('input', { type: 'radio', name: 'tab-btn', id: id, value: '' });
			if (i === 0) input.checked = true;

			var label = el('label', { 'for': id });
			label.appendChild(el('span', { 'class': 'tab-name' }, esc(cat.name)));
			label.appendChild(el('span', { 'class': 'tab-count' }, ''));

			input.addEventListener('change', function () {
				panels.forEach(function (d) { d.classList.remove('is-active'); });
				panels[i].classList.add('is-active');
			});

			tabs.appendChild(input);
			tabs.appendChild(label);
		});

		// Затем сами таблицы — все ниже строки с вкладками.
		categories.forEach(function (cat, i) {
			var panel = el('div', { id: 'content-' + (i + 1) });
			if (i === 0) panel.classList.add('is-active');
			panel.appendChild(buildTable(cat.items));
			panel.appendChild(el('p', { 'class': 'parts-empty', hidden: 'hidden' }, 'Ничего не найдено.'));

			panels.push(panel);
			tabs.appendChild(panel);
		});

		return tabs;
	}

	/* ------------------------------ поиск ------------------------------- */

	function buildSearchBox(tabs) {
		var wrap = el('div', { 'class': 'parts-search' });
		var input = el('input', {
			type: 'search',
			placeholder: 'Поиск по наименованию или коду…',
			'aria-label': 'Поиск по списку запчастей'
		});
		wrap.appendChild(input);

		input.addEventListener('input', function () {
			var q = input.value.trim().toLowerCase();

			tabs.querySelectorAll(':scope > div').forEach(function (panel, i) {
				var found = 0;

				panel.querySelectorAll('tr[data-search]').forEach(function (tr) {
					var hit = !q || tr.dataset.search.indexOf(q) !== -1;
					tr.hidden = !hit;
					if (hit) found++;
				});

				// подзаголовки разделов прячем, если в разделе ничего не осталось
				var visibleInSection = 0, lastSection = null;
				panel.querySelectorAll('tr').forEach(function (tr) {
					if (tr.dataset.section) {
						if (lastSection) lastSection.hidden = visibleInSection === 0;
						lastSection = tr; visibleInSection = 0;
					} else if (!tr.hidden && tr.dataset.search) {
						visibleInSection++;
					}
				});
				if (lastSection) lastSection.hidden = visibleInSection === 0;

				panel.querySelector('.table-wrap').hidden = (found === 0);
				panel.querySelector('.parts-empty').hidden = (found !== 0);

				var counter = tabs.querySelectorAll('label')[i].querySelector('.tab-count');
				counter.textContent = q ? ' (' + found + ')' : '';
			});
		});

		return wrap;
	}

	/* --------------------------- панель заказа -------------------------- */

	function buildOrderPanel() {
		var panel = el('div', { 'class': 'order', hidden: 'hidden' });

		panel.appendChild(el('h3', { 'class': 'order-title' }, 'Выбранные позиции'));
		panel.appendChild(el('div', { 'class': 'order-list' }));
		panel.appendChild(el('p', { 'class': 'order-total' }));
		panel.appendChild(el('p', { 'class': 'order-warning', hidden: 'hidden' }));

		var actions = el('div', { 'class': 'order-actions' });
		actions.appendChild(el('button', { type: 'button', 'class': 'btn btn-primary order-mail' }, 'Сформировать письмо'));
		actions.appendChild(el('button', { type: 'button', 'class': 'btn btn-secondary order-copy' }, 'Скопировать список'));
		actions.appendChild(el('button', { type: 'button', 'class': 'btn btn-secondary order-clear' }, 'Очистить выбор'));
		panel.appendChild(actions);

		panel.appendChild(el('p', { 'class': 'order-hint', hidden: 'hidden' }));

		return panel;
	}

	function selectedRows() {
		return allRows.filter(function (r) { return r.st.c && r.item.stock > 0; });
	}

	function orderTotal(rows) {
		return rows.reduce(function (sum, r) {
			return sum + (r.item.price === null ? 0 : r.item.price * r.st.q);
		}, 0);
	}

	function orderStamp() {
		return new Date().toLocaleString('ru-RU', {
			day: '2-digit', month: '2-digit', year: 'numeric',
			hour: '2-digit', minute: '2-digit'
		});
	}

	function contactsBlock(out) {
		out.push('');
		out.push('Мои контакты:');
		out.push('Имя: ');
		out.push('Телефон: ');
		out.push('Почта: ');
		out.push('Комментарий: ');
	}

	// Строки держим короткими: в mailto кириллица кодируется по 6 символов
	// на букву, и длинное письмо почтовый клиент может обрезать.
	function buildOrderText(rows) {
		var noPrice = rows.filter(function (r) { return r.item.price === null; });
		var out = [];

		out.push('Здравствуйте!');
		out.push('');
		out.push('Заказ с сайта ' + SITE_URL);
		out.push('Дата: ' + orderStamp());
		out.push('');
		out.push('Выбрано позиций: ' + rows.length);

		groupBy(rows, function (r) { return r.item.category; }).forEach(function (group) {
			out.push('');
			out.push('--- ' + group.name + ' ---');
			group.items.forEach(function (r, i) {
				var it = r.item;
				out.push((i + 1) + '. ' + it.name +
					(it.code ? ' (код ' + it.code + ')' : '') +
					' - ' + r.st.q + ' шт. ' +
					(it.price === null
						? '- цена по запросу'
						: 'x ' + moneyPlain(it.price) + ' = ' + moneyPlain(it.price * r.st.q)));
			});
		});

		out.push('');
		out.push('ИТОГО: ' + moneyPlain(orderTotal(rows)));
		if (noPrice.length) {
			out.push('(для ' + noPrice.length + ' поз. цена не указана, прошу уточнить)');
		}
		contactsBlock(out);

		return out.join('\n');
	}


	function mailtoURL(body) {
		var subject = 'Заказ с сайта imnaa.github.io/zipdetals от ' +
			new Date().toLocaleDateString('ru-RU');

		return 'mailto:' + MAIL_TO +
			'?cc=' + MAIL_CC +
			'&subject=' + encodeURIComponent(subject) +
			'&body=' + encodeURIComponent(body);
	}

	function copyText(text) {
		if (navigator.clipboard && navigator.clipboard.writeText) {
			return navigator.clipboard.writeText(text);
		}
		return new Promise(function (resolve, reject) {
			var ta = el('textarea');
			ta.value = text;
			ta.style.position = 'fixed';
			ta.style.opacity = '0';
			document.body.appendChild(ta);
			ta.select();
			try {
				if (document.execCommand('copy')) resolve(); else reject(new Error('copy failed'));
			} catch (e) {
				reject(e);
			} finally {
				document.body.removeChild(ta);
			}
		});
	}

	function setupOrderPanel(panel) {
		var mailBtn = panel.querySelector('.order-mail');
		var copyBtn = panel.querySelector('.order-copy');
		var clearBtn = panel.querySelector('.order-clear');
		var hint = panel.querySelector('.order-hint');

		mailBtn.addEventListener('click', function () {
			var rows = selectedRows();
			if (!rows.length) return;

			// В письмо всегда уходит полный список выбранного.
			var url = mailtoURL(buildOrderText(rows));

			// Часть почтовых программ (заметнее всего Outlook) обрезает очень
			// длинные ссылки — предупреждаем и оставляем ручной запасной путь.
			hint.hidden = url.length <= MAILTO_LIMIT;
			if (!hint.hidden) {
				hint.textContent = 'Заказ длинный. Если письмо откроется с неполным списком, ' +
					'нажмите «Скопировать список» и вставьте его в письмо — так ничего не потеряется.';
			}

			window.location.href = url;
		});

		copyBtn.addEventListener('click', function () {
			var rows = selectedRows();
			if (!rows.length) return;

			copyText(buildOrderText(rows)).then(function () {
				hint.hidden = false;
				hint.textContent = 'Список скопирован в буфер обмена.';
			}, function () {
				hint.hidden = false;
				hint.textContent = 'Не удалось скопировать автоматически — выделите текст заказа вручную.';
			});
		});

		clearBtn.addEventListener('click', function () {
			allRows.forEach(function (r) {
				r.st.c = false;
				r.st.q = 1;
				if (r.checkbox) r.checkbox.checked = false;
				if (r.input) r.input.value = 1;
			});
			saveState();
			hint.hidden = true;
			recalc();
		});
	}

	function makeRecalc(panel) {
		var list = panel.querySelector('.order-list');
		var total = panel.querySelector('.order-total');
		var warning = panel.querySelector('.order-warning');
		var title = panel.querySelector('.order-title');

		return function () {
			// суммы по строкам
			allRows.forEach(function (r) {
				var sum = r.item.price === null ? null : r.item.price * r.st.q;
				r.sumCell.innerHTML = sum === null ? '&mdash;' : esc(money(sum));
				r.tr.classList.toggle('is-picked', r.st.c);
			});

			var rows = selectedRows();
			panel.hidden = rows.length === 0;
			if (!rows.length) return;

			title.textContent = 'Выбрано позиций: ' + rows.length;

			list.innerHTML = '';
			groupBy(rows, function (r) { return r.item.category; }).forEach(function (group) {
				list.appendChild(el('h4', { 'class': 'order-group' }, esc(group.name)));

				var ul = el('ul');
				group.items.forEach(function (r) {
					var it = r.item;
					var line = esc(it.name) +
						(it.code ? ' <span class="order-code">' + esc(it.code) + '</span>' : '') +
						' &mdash; ' + r.st.q + ' шт.' +
						(it.price === null
							? ' <span class="order-code">(цена по запросу)</span>'
							: ' &times; ' + esc(money(it.price)) +
							  ' = <strong>' + esc(money(it.price * r.st.q)) + '</strong>');
					ul.appendChild(el('li', null, line));
				});
				list.appendChild(ul);
			});

			total.innerHTML = 'Итого: <strong>' + esc(money(orderTotal(rows))) + '</strong>';

			var noPrice = rows.filter(function (r) { return r.item.price === null; }).length;
			warning.hidden = noPrice === 0;
			if (noPrice) {
				warning.textContent = 'Для ' + noPrice + ' поз. цена не указана — они в сумму не вошли, ' +
					'стоимость уточняется при ответе на письмо.';
			}
		};
	}

	/* ------------------------- модальное окно фото ---------------------- */

	function setupPhotoModal(container) {
		var modal = el('div', {
			'class': 'modal fade',
			id: 'partsPhotoModal',
			tabindex: '-1',
			'aria-hidden': 'true'
		},
			'<div class="modal-dialog modal-dialog-centered modal-dialog-scrollable">' +
				'<div class="modal-content">' +
					'<div class="modal-header">' +
						'<h5 class="modal-title"></h5>' +
						'<button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Закрыть"></button>' +
					'</div>' +
					'<div class="modal-body"></div>' +
					'<div class="modal-footer">' +
						'<button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Закрыть</button>' +
					'</div>' +
				'</div>' +
			'</div>');
		document.body.appendChild(modal);

		var bsModal = new bootstrap.Modal(modal, { backdrop: 'static', keyboard: false });

		container.addEventListener('click', function (e) {
			var btn = e.target.closest('button[data-photos]');
			if (!btn) return;

			modal.querySelector('.modal-title').textContent = btn.dataset.title;
			modal.querySelector('.modal-body').innerHTML = btn.dataset.photos
				.split('|')
				.map(function (src) {
					return '<img src="' + esc(src) + '" width="100%" height="100%" loading="lazy" alt="">';
				})
				.join('<br>');

			bsModal.show();
		});
	}

	/* ------------------------------ запуск ------------------------------ */

	function showError(container, message) {
		container.innerHTML = '';
		container.appendChild(el('p', { 'class': 'parts-error' }, message));
	}

	function showNotice(container, message) {
		container.insertBefore(el('p', { 'class': 'parts-notice' }, message), container.firstChild);
	}

	// Свежая копия в обход кэша браузера; у ссылки Google уже есть «?».
	function noCache(url) {
		return url + (url.indexOf('?') === -1 ? '?' : '&') + '_=' + Date.now();
	}

	function load(url) {
		return fetch(noCache(url))
			.then(function (r) {
				if (!r.ok) throw new Error('HTTP ' + r.status);
				return r.text();
			})
			.then(function (text) {
				var rows = parseCSV(text, detectDelimiter(text));
				if (rows.length < 2) throw new Error('таблица пуста');
				return toItems(rows);
			});
	}

	function render(container, items) {
		allRows = [];
		loadState();

		var tabs = buildTabs(groupBy(items, function (it) { return it.category; }));
		var order = buildOrderPanel();

		container.innerHTML = '';
		container.appendChild(buildSearchBox(tabs));
		container.appendChild(tabs);
		container.appendChild(order);

		recalc = makeRecalc(order);
		setupOrderPanel(order);
		setupPhotoModal(container);
		recalc();
	}

	document.addEventListener('DOMContentLoaded', function () {
		var container = document.getElementById('parts');
		if (!container) return;

		load(SHEET_URL)
			.then(function (items) {
				render(container, items);
			})
			.catch(function (sheetErr) {
				// Google не ответил — показываем копию из репозитория.
				return load(BACKUP_URL)
					.then(function (items) {
						render(container, items);
						showNotice(container,
							'Google Таблица сейчас недоступна (' + esc(sheetErr.message) + '), ' +
							'показан сохранённый список — возможно, без последних изменений.');
					})
					.catch(function (backupErr) {
						showError(container,
							'Не удалось загрузить список запчастей ' +
							'(Google: ' + esc(sheetErr.message) + ', ' +
							'резервная копия: ' + esc(backupErr.message) + '). ' +
							'Если вы открыли страницу двойным щелчком по файлу — так браузер ' +
							'не даёт читать данные; откройте сайт по адресу ' +
							'<a href="https://imnaa.github.io/zipdetals/">imnaa.github.io/zipdetals</a>.');
					});
			});
	});
})();
