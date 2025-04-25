import asyncio
import logging
from re import search
import time
import pandas as pd
import argparse
from datetime import datetime, timedelta
from pathlib import Path
# from alive_progress import alive_bar, config_handler
from colorama import init, Fore
from colorama import Style

from asyncua import Client, ua

_VERSION = 0.2
_PRG_DIR = Path("./").absolute()
_LOG_FILE = _PRG_DIR / "opc2xls.log"
_MAGADAN_UTC = 11  # Магаданское время +11 часов к UTC

DEBUG = False
OPC_URL = "opc.tcp://10.100.59.1:4861"  # End point OPC
TAGS_NODE_ID = "ns=1; s=f|@LOCALMACHINE::List of all tags"  #Identifier Node Contains Tags
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


class OPC_UA_Client:
    """Simple OPC UA Client"""
    def __init__(self, endpoint: str, root_node_id: str):
        self.client = Client(endpoint)
        self.endpoint = endpoint
        self.root_node_id = root_node_id

    async def __aenter__(self):
        try:
            print(f"{Fore.WHITE}Connecting to OPC: {Fore.MAGENTA}{self.endpoint}{Style.RESET_ALL}", end="\t")
            await self.client.connect()
            print(f"{Fore.GREEN}OK {Style.RESET_ALL}")
        except:
            print(f"{Fore.RED}FAULT{Style.RESET_ALL}")
            err_msg = f"{Fore.RED}Error: Can't connect to OPC: {self.endpoint}"
            raise Exception(err_msg)
        # return self.client
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        await self.client.disconnect()

    async def get_tags(self, filter=".*") -> list:
        """Get all tags from opc and return filtered."""
        node = self.client.get_node(self.root_node_id)
        print(f"{Fore.WHITE}Get Tags", end="\t")

        # get the references of a single node
        refs = await node.get_references()
        print(f"{Fore.GREEN}OK {Style.RESET_ALL}")
        print(f"{Fore.WHITE}All Tags Items: {Fore.YELLOW}{len(refs)}{Style.RESET_ALL}")

        tags_obj = []
        tags = []
        print(f"{Fore.WHITE}Tags Filter: {Fore.YELLOW}{filter}{Style.RESET_ALL}")

        for ref in refs:
            if ref.ReferenceTypeId.Identifier == ua.ObjectIds.Organizes and search(filter, ref.BrowseName.Name) \
                    and not search(r"^@.+", ref.BrowseName.Name):
                # Organizes = 35 - Tags
                tags_obj.append(ref)

        print(f"{Fore.WHITE}Tags Items: {Fore.YELLOW}{len(tags_obj)}{Style.RESET_ALL}")

        for tag in tags_obj:
            ns = tag.NodeId.NamespaceIndex
            s = tag.NodeId.Identifier
            node = self.client.get_node(f'ns={tag.NodeId.NamespaceIndex}; s={tag.NodeId.Identifier}')

            try:
                # value = await node.read_value()
                data_value = await node.read_data_value()
                tags.append((tag.BrowseName.Name, data_value.Value.Value,
                             data_value.SourceTimestamp + timedelta(hours=_MAGADAN_UTC),))

                print(f"{Fore.WHITE}Tag:{Fore.YELLOW} {tag.BrowseName.Name}{Fore.WHITE}; "
                      f"Value:{Fore.YELLOW} {data_value.Value.Value}{Fore.WHITE}; "
                      f"TimeStamp:{Fore.YELLOW} {data_value.SourceTimestamp + timedelta(hours=_MAGADAN_UTC)}")

            except:
                print(f"{Fore.RED}Tag:{Fore.YELLOW} {tag.BrowseName.Name}{Fore.RED}; "
                      f" is BAD {Fore.RESET}")
                continue
                # value = ""
        return tags


def tags2excel(tags: list, file: Path):
    """Save list of tags to excel file"""
    df_tag = pd.DataFrame(tags, columns=["Tag", "Value", "Timestamp"])
    # Check which columns have timezones datetime64[ns, UTC]
    # df_tag.dtypes
    # Remove timezone from columns
    df_tag['Timestamp'] = df_tag['Timestamp'].dt.tz_localize(None)
    print(f"\n{Fore.WHITE}Saving to {Fore.MAGENTA}{file} {Style.RESET_ALL}", end="\t")
    try:
        with pd.ExcelWriter(file, engine="xlsxwriter") as writer:
            df_tag.to_excel(writer, sheet_name='opc2xls', startrow=2, startcol=0, index=False)
            worksheet = writer.sheets['opc2xls']
            worksheet.write(0, 0, f"OPC UA Tags from: {OPC_URL};\t Filter: {TAG_FILTER};\t Items:{df_tag.shape[0]};\t "
                                  f"DT:{datetime.now()}")
            worksheet.set_column('A:A', 27)
            worksheet.set_column('B:B', 18)
            worksheet.set_column('C:C', 19)

        print(f"{Fore.GREEN}OK {Style.RESET_ALL}")

    except Exception as e:
        print(f"{Fore.RED}FAULT{Style.RESET_ALL}")
        err_msg = f"{Fore.RED}Error: Can't write tags to: {file} {e}"
        print(err_msg)
        raise Exception(err_msg)


async def main():
    """"""
    global XLS_FILE, TAG_FILTER, OPC_URL, TAGS_NODE_ID
    start_time = time.time()
    print(f"\n{Fore.LIGHTWHITE_EX}OPC2XML v{_VERSION} - upload OPC UA tags to EXCEL{Style.RESET_ALL}")
    try:
        # Parse arguments in cmd line
        parser = argparse.ArgumentParser(
            prog=f'OPC2XLS',
            description='OPCXLS - upload OPC UA tags to EXCEL',
            epilog=f'2024 7Art v{_VERSION}'
        )

        parser.add_argument('-ep_url', type=str, default=OPC_URL,
                            help=f'OPC End Point URL (Default:"{OPC_URL}")')
        parser.add_argument('-node', type=str, default=TAGS_NODE_ID,
                            help=f'OPC NODE Identifier Contains Tags (Default: "{TAGS_NODE_ID}")')
        parser.add_argument('-filter', type=str, default=TAG_FILTER,
                            help=f'Tags Filter (Default:"{TAG_FILTER}")')
        parser.add_argument('-file', type=str, default=XLS_FILE,
                            help=f'Excel File Name (Default: "{XLS_FILE}")')
        args = parser.parse_args()

        OPC_URL = args.ep_url  # OPC and point URL
        TAGS_NODE_ID = args.node  # NODE contains PLC tags
        TAG_FILTER = args.filter  #
        XLS_FILE = args.file  #

        # Connect to OPC UA Server and get tags
        async with OPC_UA_Client(OPC_URL, TAGS_NODE_ID) as opc:
            opc_tags = await opc.get_tags(TAG_FILTER)

        # Save tags to Excel
        if len(opc_tags) > 0:
            tags2excel(opc_tags, XLS_FILE)

        print(f"{Fore.WHITE}Tags Items: {Fore.YELLOW}{len(opc_tags)}{Style.RESET_ALL}")
        print(f"{Fore.WHITE}Full time: {Fore.GREEN}{round(time.time() - start_time)} {Fore.WHITE}sec")
    except Exception as e:
        print(e)
        logger.error(e)

    finally:
        pass
        # print("Press Enter to continue...")
        # input()


if __name__ == "__main__":
    # logging.basicConfig(level=logging.INFO)
    asyncio.run(main())
