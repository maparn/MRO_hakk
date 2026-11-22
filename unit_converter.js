/* General-purpose, client-side physical unit converter.
   This tool is deliberately NOT connected to the MRO pricing engine. */
(() => {
  'use strict';

  // Third field = number of base units per 1 listed unit.
  const groups = {
    'Pikkus': { units: [
      ['mm', 'millimeeter', 0.001], ['cm', 'sentimeeter', 0.01],
      ['m', 'meeter', 1], ['km', 'kilomeeter', 1000],
      ['in', 'toll', 0.0254], ['ft', 'jalg', 0.3048],
      ['yd', 'jard', 0.9144], ['mi', 'miil', 1609.344]
    ] },
    'Mass': { units: [
      ['mg', 'milligramm', 0.000001], ['g', 'gramm', 0.001],
      ['kg', 'kilogramm', 1], ['t', 'tonn', 1000],
      ['oz', 'unts', 0.028349523125], ['lb', 'nael', 0.45359237]
    ] },
    'Maht': { units: [
      ['mL', 'milliliiter', 0.001], ['L', 'liiter', 1],
      ['m³', 'kuupmeeter', 1000],
      ['US fl oz', 'USA vedelunts', 0.0295735295625],
      ['US pt', 'USA pint', 0.473176473],
      ['US qt', 'USA quart', 0.946352946],
      ['US gal', 'USA gallon', 3.785411784],
      ['Imp gal', 'Briti gallon', 4.54609]
    ] },
    'Rõhk': { units: [
      ['Pa', 'paskal', 1], ['kPa', 'kilopaskal', 1000],
      ['bar', 'baar', 100000], ['psi', 'nael ruuttolli kohta', 6894.757293168],
      ['atm', 'standardatmosfäär', 101325]
    ] },
    'Pöördemoment': { units: [
      ['N·m', 'njuutonmeeter', 1], ['N·cm', 'njuutonsentimeeter', 0.01],
      ['lbf·ft', 'naeljalg', 1.3558179483314],
      ['lbf·in', 'naeltoll', 0.1129848290276]
    ] },
    'Temperatuur': { units: [
      ['°C', 'Celsiuse kraad', null],
      ['°F', 'Fahrenheiti kraad', null],
      ['K', 'kelvin', null]
    ] }
  };

  const category = document.getElementById('category');
  const from = document.getElementById('from');
  const to = document.getElementById('to');
  const value = document.getElementById('value');
  const result = document.getElementById('result');
  const error = document.getElementById('error');
  const note = document.getElementById('note');
  const formatter = new Intl.NumberFormat('et-EE', { maximumSignificantDigits: 10 });

  function fillUnits(select, units) {
    select.replaceChildren();
    for (const [id, label] of units) {
      const option = document.createElement('option');
      option.value = id;
      option.textContent = `${label} (${id})`;
      select.appendChild(option);
    }
  }

  function showError(message) {
    error.textContent = message;
    result.textContent = '—';
    note.textContent = '';
  }

  function temperatureToCelsius(x, unit) {
    if (unit === '°C') return x;
    if (unit === '°F') return (x - 32) * 5 / 9;
    return x - 273.15;
  }

  function celsiusToTemperature(celsius, unit) {
    if (unit === '°C') return celsius;
    if (unit === '°F') return celsius * 9 / 5 + 32;
    return celsius + 273.15;
  }

  function convert() {
    const raw = value.value.trim();
    if (raw === '' || value.validity.badInput) {
      showError('Sisesta kehtiv arv.');
      return;
    }
    const x = Number(raw);
    if (!Number.isFinite(x)) {
      showError('Sisesta lõplik arv.');
      return;
    }

    const a = from.value;
    const b = to.value;
    let y;
    if (category.value === 'Temperatuur') {
      const celsius = temperatureToCelsius(x, a);
      // Check the physical minimum in Celsius, not in the target display unit.
      // For example, -300 °F is valid, while -460 °F is below absolute zero.
      if (celsius < -273.15 - 1e-10) {
        showError('Temperatuur ei tohi olla alla absoluutse nulli (0 K).');
        return;
      }
      y = celsiusToTemperature(celsius, b);
    } else {
      const units = groups[category.value].units;
      const source = units.find(u => u[0] === a);
      const target = units.find(u => u[0] === b);
      if (!source || !target) {
        showError('Valitud ühikut ei leitud.');
        return;
      }
      y = (x * source[2]) / target[2];
    }

    if (!Number.isFinite(y)) {
      showError('Tulemus on liiga suur.');
      return;
    }
    error.textContent = '';
    result.textContent = `${formatter.format(Object.is(y, -0) ? 0 : y)} ${b}`;
    const notes = {
      'Maht': 'USA gallon (US gal) ja Briti gallon (Imp gal) on erinevad mõõtühikud.',
      'Temperatuur': 'Temperatuur teisendatakse läbi Celsiuse skaala; 0 K on absoluutne null.'
    };
    note.textContent = notes[category.value] ||
      'Füüsikaline ühikute teisendus. Lennukihoolduses kontrolli tööjuhendi nõutud ühikut ja ümardamist.';
  }

  function setupCategory() {
    const units = groups[category.value].units;
    fillUnits(from, units);
    fillUnits(to, units);
    to.selectedIndex = Math.min(1, units.length - 1);
    convert();
  }

  Object.keys(groups).forEach(name => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    category.appendChild(option);
  });

  category.value = 'Pikkus';
  setupCategory();
  category.addEventListener('change', setupCategory);
  [from, to, value].forEach(element => element.addEventListener('input', convert));
  document.getElementById('swap').addEventListener('click', () => {
    const oldFrom = from.value;
    from.value = to.value;
    to.value = oldFrom;
    convert();
  });
})();
