const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const template = fs.readFileSync(path.join(__dirname, 'templates', 'volumetry.html'), 'utf8');
const script = template.match(/<script>([\s\S]*?)<\/script>/)[1]
    .replace('const ITEMS = {{ items_json | safe }};', 'const ITEMS = testItems;')
    .replace('const BOXES = {{ boxes_json | safe }};', 'const BOXES = testBoxes;')
    .replace("{{ url_for('order_volumetry_validate', order_id=order['id']) | tojson }}", "'/validate'")
    .replace(/{% if preload_packages %}[\s\S]*?{% endif %}/, 'addPackage();');

function setup(items) {
    const elements = {
        'packages-container': {appendChild() {}},
    };
    const context = vm.createContext({
        testItems: items,
        testBoxes: [{id: 10, name: 'MB'}],
        document: {
            addEventListener() {},
            createElement() { return {}; },
            getElementById(id) { return elements[id] || null; },
        },
        setTimeout() {},
        clearTimeout() {},
    });
    vm.runInContext(script, context);
    vm.runInContext('rerenderPackage = () => {}; recalcPackage = () => {}; updateRemaining = () => {};', context);
    function repeat(id, value) {
        elements[`repeat-count-${id}`] = {value};
        elements[`repeat-error-${id}`] = {textContent: ''};
        vm.runInContext(`repeatPackage(${id})`, context);
        return elements[`repeat-error-${id}`].textContent;
    }
    return {context, repeat};
}

{
    const {context, repeat} = setup([{product_id: 1, quantity: 160, name: 'Produto'}]);
    vm.runInContext('addPackage(10, {"1": 20})', context);
    assert.equal(repeat(0, '7'), '');
    assert.equal(vm.runInContext('packages.length', context), 8);
    assert.equal(vm.runInContext('packages.reduce((sum, p) => sum + p.assignments[1], 0)', context), 160);
    assert.equal(vm.runInContext('packages.every(p => p.boxId === 10 && p.assignments[1] === 20)', context), true);
    vm.runInContext('packages[1].assignments[1] = 5', context);
    assert.equal(vm.runInContext('packages[0].assignments[1]', context), 20);
    assert.equal(vm.runInContext('packages[2].assignments[1]', context), 20);
}

{
    const {context, repeat} = setup([
        {product_id: 1, quantity: 160, name: 'Produto A'},
        {product_id: 2, quantity: 5, name: 'Produto B'},
    ]);
    vm.runInContext('addPackage(10, {"1": 20, "2": 2}); addPackage(10, {"1": 20, "2": 1})', context);
    assert.match(repeat(0, '2'), /Produto B.*até 1/);
    assert.equal(vm.runInContext('packages.length', context), 2);
    assert.equal(repeat(0, '1'), '');
    assert.equal(vm.runInContext('packages.length', context), 3);
    assert.match(repeat(0, '1'), /Produto B.*até 0/);
}

{
    const {context, repeat} = setup([{product_id: 1, quantity: 160, name: 'Produto'}]);
    vm.runInContext('addPackage()', context);
    for (const invalid of ['', '0', '-1', '1.5', 'Infinity']) {
        assert.match(repeat(0, invalid), /quantidade inteira/);
    }
    assert.match(repeat(0, '1'), /Selecione uma embalagem/);
    vm.runInContext('packages[0].boxId = 10; packages[0].boxData = testBoxes[0]', context);
    assert.match(repeat(0, '1'), /Adicione produtos/);
    assert.equal(vm.runInContext('packages.length', context), 1);
}

console.log('Volumetry repeat tests passed');
