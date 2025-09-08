class Spreadsheet {
    constructor() {
        console.log("🚀 CONSTRUCTOR: Starting with 20 rows");
        this.rows = 20;
        this.columns = 8;
        this.init();
    }

    init() {
        this.generateRows();
    }

    generateRows() {
        console.log("🔢 Generating", this.rows, "rows");
        const tbody = document.getElementById('spreadsheet-body');
        tbody.innerHTML = '';

        // Generate exactly this.rows number of rows
        for (let row = 1; row <= this.rows; row++) {
            const tr = document.createElement('tr');
            tr.className = 'data-row';
            
            // Row header
            const th = document.createElement('th');
            th.className = 'row-header';
            th.textContent = row;
            tr.appendChild(th);

            // Data cells
            for (let col = 0; col < this.columns; col++) {
                const td = document.createElement('td');
                td.className = 'cell';
                
                const input = document.createElement('input');
                input.type = 'text';
                input.value = '';
                
                td.appendChild(input);
                tr.appendChild(td);
            }

            tbody.appendChild(tr);
        }
        
        console.log("✅ Generated", tbody.children.length, "rows in DOM");
    }
}

// Initialize when page loads
document.addEventListener('DOMContentLoaded', () => {
    console.log("🌟 Page loaded, creating spreadsheet");
    window.spreadsheet = new Spreadsheet();
});
