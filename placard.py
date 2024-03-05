import openpyxl
import barcode
from barcode.writer import ImageWriter
import re
import os

def print_html_barcode(content: str):
    with open('index.html', 'r', encoding='utf-8') as file:
        html_content = file.read()

        modified_html = html_content.replace('<body>','<body>'+content)
 
        # Write the modified HTML content to a new file
        with open('index_print.html', 'w', encoding='utf-8') as file:
            file.write(modified_html)

def create_html_barcode(products: dict) -> str:
    content = ''
    for i, key in enumerate(products):
        if i % 2 == 0:
            content += '<div class="row my-5">'
        
        content += f'''
            <div class="col-6">
                <div class="product-card">
                    <div class="product-info">
                        <div class="product-name"><b>{products[key]['name']}</b></div>
                        <div class="product-code">{products[key]['code']}</div>
                    </div>
                    <div class="below">
                        <div class="product-image">
                            <img alt='{products[key]['code']}'
                                src='https://barcode.tec-it.com/barcode.ashx?data={products[key]['code']}%0a&code=Code128&translate-esc=on'/>
                        </div>
                        <div class="product-price">
                            <div class="currency"><b>Rp.</b></div>
                            <span>
                                <span class="amount">{format_number(products[key]['price'])}</span>
                                <span class="unit"> /{products[key]['unit']}</span>
                            </span>
                        </div>
                    </div>
                </div>
            </div>
        '''

        if i % 2 == 1:
            content += '</div>'

    if len(products) % 2 == 1:
            content += '</div>'

    return content

def pretty_print(products: dict):
    if len(products) > 0:
        print('==============================')
        print('PREVIEW DATA PRODUK')
        for co, key in enumerate(products, start=1):
            print('==============================')
            print('#', co)
            print(f'Name: {products[key]["name"]}')
            print(f'Price: {products[key]["price"]}')
            print(f'Code: {products[key]["code"]}')
            print(f'Unit: {products[key]["unit"]}')
        print('==============================')

def format_number(num):
    # Convert to string and remove non-digit characters
    num_str = re.sub(r'\D', '', str(num))
    # Insert '.' every three digits from the end
    formatted_num = re.sub(r'(\d)(?=(\d{3})+$)', r'\1.', num_str)
    return formatted_num

def main():
    content = ''
    products = {}

    while(1):
        os.system('cls' if os.name == 'nt' else 'clear')
        pretty_print(products)
        print('Banyak Data: ', len(products))
        print('1. Upload Data Dari Excel')
        print('2. Buat Barcode')
        print('3. Pilih Data Yang Ingin Di Pilih')
        print('4. Keluar')
        choice = int(input('Pilih: '))

        if choice == 1:
            excel_name = input('Masukkan Nama File Excel yang ingin di buat barcode: ')
            # excel_sheet = input('Masukkan Nama Sheet Excel yang ingin di buat barcode: ')

            workbook = openpyxl.load_workbook(excel_name)
            sheet = workbook.active
            
            for co, row in enumerate(sheet.iter_rows(values_only=True), start=1):
                if co == 1:
                    continue
                
                if row[0] is None or row[1] is None or row[2] is None or row[3] is None:
                    print('Data di excel di baris ke', co, 'tidak lengkap')
                    continue
                
                products[str(row[0])] = {
                    'name': row[1],
                    'price': row[2],
                    'code': row[0],
                    'unit': row[3],
                }

                code39 = barcode.codex.Code39(str(row[0]), writer=ImageWriter(), add_checksum=False)
                code39.save('img/' + str(row[0]))

            content = create_html_barcode(products)
            
            input('Data berhasil di upload, tekan enter untuk melanjutkan')

        elif choice == 2:
            print_html_barcode(content)
            input('Barcode berhasil di buat, tekan enter untuk melanjutkan')

        elif choice == 3:
            selected = input('Pilih kode produk yang ingin di buat barcode dipisahkan spasi (contoh: 12 123 1234): ').split(' ')
            selected_products = {}
            for key in selected:
                if key in products: selected_products[key] = products[key]
                else: print('Kode produk', key, 'tidak ditemukan')

            print_html_barcode(create_html_barcode(selected_products))
            input('Data berhasil di upload, tekan enter untuk melanjutkan')

        elif choice == 4: break
        else: break

if __name__ == '__main__':
    main()
    print('Berhasil membuat barcode')