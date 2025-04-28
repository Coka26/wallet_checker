import requests
from web3 import Web3
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill

# Список сетей и RPC. Беспланые можно взять на сайте https://dashboard.alchemy.com/apps или у GPT
chains = [
    { "name": "OP Mainnet", "rpc": "" },
    { "name": "Mode", "rpc": "" },
    { "name": "Unichain", "rpc": "" },
    { "name": "Lisk", "rpc": "" },
    { "name": "Soneium", "rpc": "" },
    { "name": "Ink", "rpc": "" },
    { "name": "Base", "rpc": "" }
]

# Список твоих кошельков
wallets = [
    "0x516d5ce74c65C10908b233Fc81827E8Dbe84E40a",
    "0xe221933c6EF05B0B20106F138748282b3016ac2D",
  

    # добавь остальные кошельки сюда
]

def wei_to_eth(wei):
    return round(Web3.from_wei(wei, 'ether'), 6)  # Округляем до 6 знаков

def get_wallet_info(rpc_url, wallet):
    try:
        web3 = Web3(Web3.HTTPProvider(rpc_url))
        if not web3.is_connected():
            return None
        
        wallet = Web3.to_checksum_address(wallet)  # <-- ДОбавил тут!
        
        balance = web3.eth.get_balance(wallet)
        tx_count = web3.eth.get_transaction_count(wallet)
        latest_block = web3.eth.block_number
        
        return {
            "balance": wei_to_eth(balance),
            "transactions": tx_count,
            "latest_block": latest_block
        }
    except Exception as e:
        print(f"Ошибка: {e}")
        return None


def save_to_excel(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Wallet Info"

    headers = ["Сеть", "Кошелек", "Баланс (ETH)", "Количество транзакций", "Последний блок"]
    ws.append(headers)

    green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Светло-зеленый
    red_fill = PatternFill(start_color="FF7F7F", end_color="FF7F7F", fill_type="solid")    # Светло-красный
    orange_fill = PatternFill(start_color="FFD580", end_color="FFD580", fill_type="solid") # Светло-оранжевый

    for item in data:
        ws.append([item['chain'], item['wallet'], item['balance'], item['transactions'], item['latest_block']])
        row_idx = ws.max_row

        # Закрашиваем баланс
        ws.cell(row=row_idx, column=3).fill = green_fill

        # Закрашиваем количество транзакций в зависимости от диапазона
        tx_cell = ws.cell(row=row_idx, column=4)
        if item['transactions'] <= 50:
            tx_cell.fill = red_fill
        elif item['transactions'] <= 99:
            tx_cell.fill = orange_fill
        else:
            tx_cell.fill = green_fill

    # Настраиваем ширину колонок
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[column_letter].width = max_length + 2

    filename = f"wallet_info_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    wb.save(filename)
    print(f"\nДанные сохранены в файл: {filename}")

def main():
    results = []
    for chain in chains:
        print(f"\n=== Сеть: {chain['name']} ===\n")
        for wallet in wallets:
            info = get_wallet_info(chain['rpc'], wallet)
            if info:
                print(f"Кошелек: {wallet}")
                print(f"Баланс: {info['balance']} ETH")
                print(f"Количество транзакций: {info['transactions']}")

                print("-" * 50)
                results.append({
                    "chain": chain['name'],
                    "wallet": wallet,
                    "balance": info['balance'],
                    "transactions": info['transactions'],
                    "latest_block": info['latest_block']
                })
            else:
                print(f"Не удалось подключиться к сети {chain['name']} для кошелька {wallet}")
    save_to_excel(results)

if __name__ == "__main__":
    main()
    
# библиотека openpyxl (для создания Excel-файла).
# Установи её, если ещё не установил:  pip install openpyxl
# библиотека colorama, установи её через команду:   pip install colorama
# Для запуска python wallet_checker.py