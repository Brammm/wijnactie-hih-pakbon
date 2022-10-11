if (!'content' in document.createElement('template')) {
    alert('Gebruik een andere browser');
}

const formatter = new Intl.NumberFormat('nl-BE', {
    style: 'currency',
    currency: 'EUR',
})

const config = [
    {
        "type": "ROOD",
        "name": "Paul Mas - Cabernet Sauvignon - 2020",
        "price": 7.10,
    },
    {
        "type": "ROOD",
        "name": "La Forge Estate - 2020",
        "price": 9.65,
    },
    {
        "type": "ROOD",
        "name": "Acanto - IT - 2021",
        "price": 10.40,
    },
    {
        "type": "ROOD",
        "name": "Plan Vermeersch - FR - 2021",
        "price": 12.30,
    },
    {
        "type": "ROOD",
        "name": "Tour Baladoz - FR - 2020",
        "price": 28.70,
    },
    {
        "type": "ROSE",
        "name": "Paul Mas - ROSE - FR - 2021",
        "price": 7.79,
    },
    {
        "type": "WIT",
        "name": "Valmont WIT - FR - 2021",
        "price": 6.99,
    },
    {
        "type": "WIT",
        "name": "Vignes de Nicole - FR - 2021",
        "price": 10.55,
    },
    {
        "type": "WIT",
        "name": "A Teilleira Parcelas - ESP - 2021",
        "price": 11.35,
    },
    {
        "type": "WIT",
        "name": "Sancerre - FR - 2021",
        "price": 19.29,
    },
    {
        "type": "BUBBELS",
        "name": "CAVA",
        "price": 9.19,
    },
    {
        "type": "PORTO",
        "name": "Porto rood",
        "price": 8.75,
    },
];

function write(element, selector, content) {
    element.querySelector(selector).textContent = content;
}

var input = document.getElementById('input')
input.addEventListener('change', function () {
    readXlsxFile(input.files[0]).then(function (rows) {
        rows.shift(); // remove first row with the headers

        // Instantiate the table with the existing HTML tbody
        // and the row with the template
        const orders = document.querySelector('#orders');
        const tableTemplate = document.querySelector('#table');
        const rowTemplate = document.querySelector('#row');
        rows.forEach((orderRow) => {
            // Clone the new row and insert it into the table
            const table = tableTemplate.content.cloneNode(true);
            const tbody = table.querySelector('tbody')

            let totalBottles = 0;
            let totalPrice = 0;

            config.forEach((wine, index) => {
                const row = rowTemplate.content.cloneNode(true);
                const bottleCount = orderRow[index + 2];
                const tr = row.querySelector('tr');

                totalBottles += bottleCount;
                totalPrice += wine.price * bottleCount;

                tr.cells[0].textContent = index + 1;
                tr.cells[1].textContent = wine.type;
                tr.cells[2].textContent = wine.name;
                tr.cells[3].innerHTML = formatter.format(wine.price);
                if (bottleCount > 0) {
                    tr.cells[4].textContent = formatter.format(wine.price * bottleCount);
                    tr.cells[5].textContent = bottleCount;
                }

                tbody.appendChild(row);
            });

            write(table, '.name', orderRow[21] || orderRow[1]);
            write(table, '.contact', orderRow[20]);
            write(table, '.total_bottles', totalBottles);
            write(table, '.total_price', formatter.format(totalPrice));
            write(table, '.payment', orderRow[14]);
            write(table, '.delivery', orderRow[16]);

            orders.appendChild(table);
        });
    })
})
