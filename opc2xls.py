import asyncio
import logging
from re import search
import time
import pandas as pd  # Set -ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted
import argparse
from datetime import datetime, timedelta
from pathlib import Path
# from alive_progress import alive_bar, config_handler
from colorama import init, Fore
from colorama import Style

from asyncua import Client, ua

_VERSION = 0.1
_PRG_DIR = Path("./").absolute()
_LOG_FILE = _PRG_DIR / "opc2xls.log"
_MAGADAN_UTC = 11  # Магаданское время +11 часов к UTC

DEBUG = False
OPC_URL = "opc.tcp://10.100.59.1:4861"
XLS_FILE = _PRG_DIR / "opc2xls.xlsx"
TAG_FILTER = r''

init(autoreset=True)
logger = logging.getLogger()
log_format = f"%(asctime)s - %(levelname)s -(%(funcName)s(%(lineno)d) - %(message)s"
logging.basicConfig(
    format=log_format,
    level=logging.DEBUG if DEBUG else logging.ERROR,
    # filename=_LOG_FILE,
)


class OPC_MH_Client:
    def __init__(self, endpoint):
        self.client = Client(endpoint)

    async def __aenter__(self):
        try:
            await self.client.connect()
        except:
            err_msg = f"{Fore.RED}Error: Can't connect to OPC"
            raise Exception(err_msg)
        return self.client

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        await self.client.disconnect()

    async def get_tags(self, match=".*"):
        pass


async def get_opc_tags(ep_url: str, tag_filter: str) -> list:
    """Connect to OPC UA Server and get all tags"""
    print(f"{Fore.WHITE}Connecting to OPC {Fore.YELLOW}{ep_url}{Style.RESET_ALL}", end="\t")
    async with (OPC_MH_Client(OPC_URL) as client):
        print(f"{Fore.GREEN}OK {Style.RESET_ALL}", )
        obj = client.get_node('ns=1; s=f|@LOCALMACHINE::List of all tags')
        print(f"{Fore.WHITE}Getting Tags", end="\t")
        # get the references of a single node
        refs = await obj.get_references()
        print(f"{Fore.GREEN}OK {Style.RESET_ALL}")
        print(f"{Fore.WHITE}All Tags Items: {Fore.YELLOW}{len(refs)}{Style.RESET_ALL}")
        tags_obj = []
        tags = []
        print(f"{Fore.WHITE}Tags Filter: {Fore.YELLOW}{tag_filter}{Style.RESET_ALL}")

        for ref in refs:
            if ref.ReferenceTypeId.Identifier == ua.ObjectIds.Organizes and search(tag_filter, ref.BrowseName.Name) \
                    and not search(r"^@.+", ref.BrowseName.Name):
                # Organizes = 35 - Tags
                tags_obj.append(ref)

        print(f"{Fore.WHITE}Tags Items: {Fore.YELLOW}{len(tags_obj)}{Style.RESET_ALL}")

        for tag in tags_obj:
            ns = tag.NodeId.NamespaceIndex
            s = tag.NodeId.Identifier
            node = client.get_node(f'ns={tag.NodeId.NamespaceIndex}; s={tag.NodeId.Identifier}')

            try:
                # value = await node.read_value()
                data_value = await node.read_data_value()
                tags.append((tag.BrowseName.Name, data_value.SourceTimestamp + timedelta(hours=_MAGADAN_UTC),
                             data_value.Value.Value))

                print(f"{Fore.WHITE}Tag:{Fore.YELLOW} {tag.BrowseName.Name}{Fore.WHITE}; "
                      f"Value:{Fore.YELLOW} {data_value.Value.Value}{Fore.WHITE}; "
                      f"TimeStamp:{Fore.YELLOW} {data_value.SourceTimestamp + timedelta(hours=_MAGADAN_UTC)}")

            except:
                continue
                # value = ""
        return tags


def tags2excel(tags: list, excel_file: Path):
    """Saves tags to excel file"""
    df_tag = pd.DataFrame(tags, columns=["Tag", "Timestamp", "Value"])
    # Check which columns have timezones datetime64[ns, UTC]
    # df_tag.dtypes
    # Remove timezone from columns
    df_tag['Timestamp'] = df_tag['Timestamp'].dt.tz_localize(None)
    with pd.ExcelWriter(excel_file, engine="xlsxwriter") as writer:
        df_tag.to_excel(writer, sheet_name='opc2xls', startrow=2, startcol=0, index=False)
        worksheet = writer.sheets['opc2xls']
        worksheet.write(0, 0, f"OPC Tags from: {OPC_URL}; Filter: {TAG_FILTER}; Items:{df_tag.shape[0]}; "
                              f"DT:{datetime.now()}")



async def main():
    """"""
    global XLS_FILE, TAG_FILTER, OPC_URL
    start_time = time.time()
    parser = argparse.ArgumentParser(
        prog=f'OPC2XLS',
        description='OPCXLS -  Utility export OPC UA tags to Excel files',
        epilog=f'2024 7Art v{_VERSION}'
    )
    try:
        parser.add_argument('-ep_url', type=str, default=OPC_URL,
                            help=f'OPC End Point URL (Default:{OPC_URL})')
        parser.add_argument('-filter', type=str, default=TAG_FILTER,
                            help=f'Tags filter (Default:{TAG_FILTER})')
        parser.add_argument('-file_name', type=str, default=XLS_FILE,
                            help=f'excel file name (Default:{XLS_FILE})')
        args = parser.parse_args()
        print(f"\n{Fore.LIGHTWHITE_EX}OPC2XML v{_VERSION} - uploads OPC UA tags to EXCEL{Style.RESET_ALL}")
        OPC_URL = args.ep_url
        TAG_FILTER = args.filter
        XLS_FILE = args.file_name

        opc_tags = await get_opc_tags(args.ep_url, args.filter)
        # Save tags to Excel
        print(f"{Fore.WHITE}Saving to {Fore.MAGENTA}{XLS_FILE} {Style.RESET_ALL}", end="\t")
        tags2excel(opc_tags, args.file_name)
        print(f"{Fore.GREEN}OK {Style.RESET_ALL}")

        print(f"{Fore.WHITE}Full time: {Fore.GREEN}{round(time.time() - start_time)} {Fore.WHITE}sec;")
    except Exception as e:
        print(e)
        logger.error(e)

    finally:
        pass
        # print("Press Enter to continue...")
        # input()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    asyncio.run(main())
