
import os
from kiteconnect import KiteConnect
import time, json, datetime, sys
import xlwings as xw
import pandas as pd

def get_login_credentials():
	global login_credential

	def login_credentials():
                print("---- Enter you Zerodha login_credentials ---")
                login_credential = {"api_key": str(input("Enter API Key :")),
                                "api_secret": str(input("Enter API Secret :"))
                                }
                if input("Press Y to save login_credential and any key to bypass : ").upper() == "Y":
                        with open(f"login_credentials.txt", "w") as f:
                                json.dump(login_credential, f)
                        print("Data Saved. ..")
                else:
                        print("Data save canceled!!!")

        while True:
                try:
                        with open(f"login_credentials.txt", "r") as f:
                                login_credential = json.load(f)
                        break
                except:
                        login_credentials()
        return login_credential

def get_access_token():
        global access_token

        def login():
                global login_credential
                print("Trying Log In...")
                kite = KiteConnect(api_key=login_credential["api_key"])
                print("Login url : ", kite.login_url())
                request_tkn = input("Login and enter your request token here : ")
                try:
                        access_token = kite.generate_session(request_token=request_tkn, api_secret=login_credential["api_secret"])['access _token']
                        os.makedirs(f"AccessToken", exist_ok=True)
                        with open(f"AccessToken/{datetime.datetime.now().date()}. json", "w") as f:
                                json.dump (access_token, f)
                        print("Login successful...")
                except Exception as e:
                        print(f"Login Failed {{{e}}}")

        print("Loading access token...")
        while True:
                if os.path.exists(f"AccessToken/{datetime.datetime.now().date()}.json"):
                        with open(f"AccessToken/{datetime.datetime.now().date()}.json", "r") as f:
                                access_token = json.load(f)
                        break
                else:
                        login()
        return access_token

def get_kite():
        global kite, login_credential, access_token
        try:
                kite = KiteConnect(api_key=login_credential["api_key"])
                kite.set_access_token(access_token)
        except Exception as e:
                print(f"Error : {e}")
                os.remove(f"AccessToken/{datetime.datetime.now().date()}. json") if os.path.exists(
                        f"AccessToken/{datetime.datetime.now().date()}.json") else None
                sys.exit()

def get_live_data(instruments):
        global kite, live_data
        try:
                live_data
        except:
                live_data = {}
        try:
                live_data = kite.quote(instruments)
        except Exception as e:
                # print(f'Get live_data Failed {{{e}}}")
                pass
        return live_data

def place_trade(symbol, quantity, direction):
        try:
                order = kite.place_order(variety=kite.VARIETY_REGULAR,
                        exchange=symbol[0:3],
                        tradingsymbol=symbol[4:],
                        transaction_type=kite.TRANSACTION_TYPE_BUY if direction == "Buy" else kite.TRANSACTION_TYPE_SELL,
                        quantity=int(quantity),
                        product=kite.PRODUCT_MIS,
                        order_type=kite.ORDER_TYPE_MARKET,
                        price=0.0,
                        validity=kite.VALIDITY_DAY,
                        tag="TA Python")
                # print(f"symbol {symbol}, Qty {quantity}, Direction {direction}, Time {datetime.datetime.now().time()} \n"
                #       f"      => {order}")
                return order
        except Exception as e:
                return f"{e}"

def get_orderbook():
        global orders
        try:
                orders
        except:
                orders = {}
        try:
                data = pd.DataFrame(kite.orders())
                data = data[data["tag"] == "TA Python"]
                data = data.filter(["order_timestamp", "exchange", "tradingsymbol", "transaction_type", "quantity", "average price", "status", "status_message_raw"])
                data.columns = data.columns.str.replace("_", " ")
                data.columns = data.columns.str.title()
                data = data.set_index(["Order Timestamp"], drop=True)
                data = data.sort_index(ascending=True)
                orders = data
        except Exception as e:
                # print(e)
                pass
        return orders

def start_excel():
        global kite, live_data
        print("Excel Starting...")
        if not os.path.exists("TA_Python.xlsx"):
                try:
                        wb = xw.Book()
                        wb.save('TA_Python.xlsx')
                        wb.close()
                except Exception as e:
                        print(f"Error : {e}")
                        sys.exit()
        wb = xw.Book("TA_Python.xlsx")
        for i in ["Data", "Exchange", "OrderBook"]:
                try:
                        wb.sheets(i)
                except:
                        wb.sheets.add(i)        
        dt = wb.sheets("Data")
        ex = wb.sheets("Exchange")
        ob = wb.sheets("OrderBook")
        ex.range("a:j").value = ob.range("a:h").value =  dt.range("p:q").value = None
        dt.range(f"a1:q1").value = ["Sr No", "Symbol", "Open", "High", "Low", "LTP", "Volume", "Vwap", "Best Bid Price", "Best Ask Price", "Close", "Qty", "Direction", "Entry Signal", "Exit Signal", "Entry", "Exit"]

        subs_lst = []
        while True:
                try:
                        master_contract = pd.DataFrame(kite.instruments())
                        master_contract = master_contract.drop(["instrument_token","exchange_token","last_price","tick_size"],axis=1)
                        master_contract["watchlist_symbol"] = master_contract["exchange"] + ":" + master_contract["tradingsymbol"]
                        master_contract.columns = master_contract.columns.str.replace("_", " ")
                        master_contract.columns = master_contract.columns.str.title()
                        ex.range("a1").value = master_contract
                        break
                except Exception as e:
                        time.sleep(1)
        while True:
                try:
                        time.sleep(0.5)
                        get_live_data(subs_lst)
                        symbols = dt.range(f"b{2}:b{500}").value
                        trading_info = dt.range(f"l{2}:q{500}").value

                        for i in subs_lst:
                                if i not in symbols:
                                        subs_lst.remove(i)
                                        try:
                                                del live_data[i]
                                        except Exception as e:
                                                pass
                        main_list = []
                        idx = 0
                        for i in symbols:
                                lst = [None, None, None, None, None, None, None, None, None]
                                if i:
                                        if i not in subs_lst:
                                                subs_lst.append(i)
                                        if i in subs_lst:
                                                try:
                                                        lst = [live_data[i]["ohlc"]["open"],
                                                                live_data[i]["ohlc"]["high"],
                                                                live_data[i]["ohlc"]["low"],
                                                                live_data[i]["last_price"]]
                                                        
                                                        try:
                                                                lst += [live_data[i]["volume"],
                                                                live_data[i]["average_price"],
                                                                live_data[i]["depth"]["buy"][0]["price"],
                                                                live_data[i]["depth"]["sell"][0]["price"],
                                                                live_data[i]["ohlc"]["close"]]
                                                        except:
                                                                lst += [0, 0, 0, 0, live_data[i]["ohlc"]["close"]]
                                                        trade_info = trading_info[idx]
                                                        if trade_info[0] is not None and trade_info[1] is not None:
                                                                if type(trade_info[0]) is float and type(trade_info[1]) is str:
                                                                        if trade_info[1].upper() == "BUY" and trade_info[2] is True:
                                                                                if trade_info[2] is True and trade_info[3] is not True and trade_info[4] is None and trade_info[5] is None:
                                                                                        dt.range(f"p{idx + 2}").value = place_trade(i, int(trade_info[0]), "Buy")
                                                                                elif trade_info[2] is True and trade_info[3] is True and trade_info[4] is not None and trade_info[5] is None:
                                                                                        dt.range(f"q{idx + 2}").value = place_trade(i, int(trade_info[0]), "sell")
                                                                        if trade_info[1].upper() == "SELL" and trade_info[2] is True:
                                                                                if trade_info[2] is True and trade_info[3] is not True and trade_info[4] is None and trade_info[5] is None:
                                                                                        dt.range(f"p{idx + 2}").value = place_trade(i, int(trade_info[0]), "Sell")
                                                                                elif trade_info[2] is True and trade_info[3] is True and trade_info[4] is not None and trade_info[5] is None:
                                                                                        dt.range(f"q{idx + 2}").value = place_trade(i, int(trade_info[0]), "Buy")
                                                except Exception as e:
                                                        print(e)
                                                        pass
                                main_list.append(lst)
                                idx += 1

                        dt.range("c2").value = main_list
                        if wb.sheets.active.name == "OrderBook":
                                ob.range("a1").value = get_orderbook()
                except Exception as e:
                        # print(e)
                        pass
# subscribe "TA Python" on YouTube

if __name__ == '__main__':
        # get_login_credentials()
        # get_access_token()
        # get_kite()
        # get_orderbook()
        # start_excel()
        print("Hellow world")
                