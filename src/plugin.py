import ctypes

try:
    # to fix blurry text on Windows
    # https://stackoverflow.com/questions/50884283/
    ctypes.windll.shcore.SetProcessDpiAwareness(True)
except:
    pass

import os
import pcbnew
import wx
from datetime import datetime
import wx.lib.mixins.listctrl as listmix
from .ki_result_event import EVT_RESULT
from .ki_push_thread import PushThread
from .utils import get_symbol_dict, \
    pcb_2_sch_path, get_sch_file_name, json_from_bom__with_pn_as_key

class EditableListCtrl(wx.ListCtrl, listmix.TextEditMixin):
    """ TextEditMixin allows any column to be edited. """
    def __init__(self, parent, _id=wx.ID_ANY, pos=wx.DefaultPosition,
                 size=wx.DefaultSize, style=0):
        # use `_id` to avoid built-in `id`
        wx.ListCtrl.__init__(self, parent, _id, pos, size, style)
        listmix.TextEditMixin.__init__(self)


# no code outside of class, or the plugin will not show
class BOMFrame(wx.Frame):
    def __init__(self, parent=None, title='Save parts to Digi-Key myLists'):
        super(BOMFrame, self).__init__(parent, title=title, size=(1000, 600))

        # prevent user from making the frame too small
        self.SetSizeHints(960, 610)

        # data
        self.bom = {}
        self.bom_by_pn_field = {}
        self.max_list_length = 1_000_000

        board = pcbnew.GetBoard()
        pcb_path = board.GetFileName()
        self.kicad_sch_path = pcb_2_sch_path(pcb_path)
        self.list_name = datetime.today().strftime('%Y%m%d') + "_" + get_sch_file_name(self.kicad_sch_path)
        self.wx_md = None

        try:
            self.symbol_dict = get_symbol_dict(self.kicad_sch_path)
        except FileNotFoundError:
            error_caption = 'Schematic file (.kicad_sch) not found'
            error_message = \
                'The "Push to Digi-Key myLists" plugin requires KiCad schematic file (.kicad_sch).\n\n' \
                'If you have an old schematic file (.sch), save it with the Schematic Editor. ' \
                'This will create the equivalent KiCad schematic file (.kicad_sch).\n\n' \
                'The schematic file (.kicad_sch) and the PCB file (.kicad_pcb) must have the same name.'.format()
            self.show_error_message_then_exit(error_message, error_caption)
        except:
            error_caption = 'Error parsing schematic file'
            error_message = \
                'There was an error parsing \n{path}.\n\n' \
                'If you have an old schematic file (.sch), save it with the Schematic Editor. ' \
                'This will create the equivalent KiCad schematic file (.kicad_sch).\n\n' \
                'The schematic file (.kicad_sch) and the PCB file (.kicad_pcb) must have the same name.' \
                .format(path=self.kicad_sch_path)
            self.show_error_message_then_exit(error_message, error_caption)

        self.current_pn_field_str = "Partno"
        # layout
        self.panel = wx.Panel(self)
        self.gbs = wx.GridBagSizer(0, 0)
        self.gbs_inputs = wx.GridBagSizer(0, 0)
        self.gbs_input_pn_field = wx.GridBagSizer(0, 0)
        self.gbs_input_list_name = wx.GridBagSizer(0, 0)
        self.gbs_inputs = wx.GridBagSizer(0, 0)
        self.gbs_button = wx.GridBagSizer(0, 0)
        self.gbs_progress = wx.GridBagSizer(0, 0)
        self.wx_pn_field_question = None
        self.wx_pn_field_dropdown = None
        self.wx_list_name_label = None
        self.wx_list_name_input = None
        self.wx_progress_gauge = None
        self.wx_progress_text = None
        self.wx_push_btn = None
        self.wx_dummy_text_1 = None
        self.wx_tos_text = None
        self.wx_bom_lc = EditableListCtrl(self.panel, size=(970, 480), style=wx.LC_REPORT | wx.LC_HRULES)
        self.InitUI()
        self.Centre()
        self.Show()

    def InitUI(self):
        self.update_listctrl_with_qty(self.symbol_dict, "Partno")
        self.update_bom_by_pn_field(self.current_pn_field_str)

        self.wx_list_name_label = wx.StaticText(self.panel, label='List name')
        self.wx_list_name_input = wx.TextCtrl(self.panel, value=self.list_name)
        self.wx_list_name_input.Bind(wx.EVT_TEXT, self.on_list_name_change)

        self.wx_push_btn = wx.Button(self.panel, label='Create DigiKey List', size=(200, 40))
        self.wx_push_btn.Bind(wx.EVT_BUTTON, self.post_bom_data)

        self.wx_dummy_text_1 = wx.StaticText(self.panel, label='           ', size=(200, 2))
        self.wx_progress_text = wx.StaticText(self.panel, label='PROGRESS TEXT', size=(200, 25))
        self.wx_progress_gauge = wx.Gauge(self.panel, range=100, size=(200, 12), style=wx.GA_HORIZONTAL)

        self.gbs_inputs.Add(self.wx_list_name_label, pos=(0, 0), span=(1, 1),
                            flag=wx.EXPAND | wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.gbs_inputs.Add(self.wx_list_name_input, pos=(1, 0), span=(1, 1),
                            flag=wx.EXPAND | wx.ALL | wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.gbs_button.Add(self.wx_push_btn, 
                            flag=wx.EXPAND | wx.ALL | wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, pos=(0, 2), span=(1, 1))

        self.gbs_progress.Add(self.wx_progress_text, pos=(1, 0), span=(1, 1))
        self.gbs_progress.Add(self.wx_progress_gauge, pos=(2, 0), span=(1, 1))

        # Main Grid
        self.gbs.Add(self.gbs_inputs, pos=(0, 0), span=(1, 1))
        self.gbs.Add(self.gbs_button, pos=(0, 1), span=(1, 1),
                     flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.gbs.Add(self.gbs_progress, pos=(0, 2), span=(1, 1),
                     flag=wx.ALIGN_LEFT | wx.LEFT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.gbs.Add(self.wx_bom_lc, pos=(1, 0), span=(1, 3),
                     flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=5)

        # Allow the listctrl to fill the window if the window grows.
        # Usually the row_idx and col_idx of the listctrl in the GridBagSizer.
        # Be careful, this might mess with some widget's alignment.
        self.gbs.AddGrowableRow(1)
        self.gbs.AddGrowableCol(2)
        self.panel.SetSizerAndFit(self.gbs)

        # Hide controls after `SetSizerAndFit()`
        self.wx_progress_text.Hide()
        self.wx_progress_gauge.Hide()

        EVT_RESULT(self, self.message_handler)  # sync state between Push Thread and the UI
   
    def update_listctrl_with_qty(self, symbol_dict, pn_field: str):
        """ update the `self.wx_bom_lc` """

        # header
        self.wx_bom_lc.ClearAll()
        self.wx_bom_lc.InsertColumn(0, 'Row', wx.LIST_FORMAT_LEFT, 75)  # numbering the rows for better navigation
        self.wx_bom_lc.InsertColumn(1, 'Part Number*', wx.LIST_FORMAT_LEFT, 200)  # *: mandatory
        self.wx_bom_lc.InsertColumn(2, 'Quantity*', wx.LIST_FORMAT_LEFT, 75)
        self.wx_bom_lc.InsertColumn(3, 'Customer Reference', wx.LIST_FORMAT_LEFT, 600)

        symbol_dict = dict(sorted(symbol_dict.items()))

        # rows
        _idx = 0
        for _idx, symbol_pn in enumerate(symbol_dict):
            row_index = self.wx_bom_lc.InsertItem(self.max_list_length, _idx)
            self.wx_bom_lc.SetItem(row_index, 0, str(_idx + 1))
            self.wx_bom_lc.SetItem(row_index, 1, str(symbol_pn))
            self.wx_bom_lc.SetItem(row_index, 2, str(symbol_dict[symbol_pn]['Quantity']))
            self.wx_bom_lc.SetItem(row_index, 3, str(symbol_dict[symbol_pn]['References']))

    def update_bom_by_pn_field(self, pn_field: str):
        self.bom_by_pn_field[pn_field] = self.bom
        
    def on_list_name_change(self, event):
        _str_value = self.wx_list_name_input.GetValue()
        self.list_name = _str_value.strip()

    def message_handler(self, message):
        _data = message.data
        if _data['state'] == 'Finished':
            self.Close()
        elif _data['state'] == 'ERR_REQUESTS_EXCEPTION':
            error_caption = 'Cannot push part data to Digi-Key myLists'
            error_message = \
                'The plugin cannot push part data to Digi-Key myLists.\n\n' \
                'Cannot connect to {url}.\n\n' \
                'If your computer is behind a proxy system, please contact your administrator.' \
                ''.format(url=_data['api_url'])
            self.show_error_message_then_exit(message=error_message, caption=error_caption)
        elif _data['state'] == 'ERR_SENDING_REQUEST':
            error_caption = 'Cannot push part data to Digi-Key myLists'
            error_message = 'There was an error pushing part data to Digi-Key myLists.'
            self.show_error_message_then_exit(message=error_message, caption=error_caption)
        elif _data['state'] == 'SHORT_URL_NOT_RETURNED':
            error_caption = 'Cannot create Digi-Key list'
            error_message = \
                'There was an error while creating a Digi-Key list from the given part numbers.\n' \
                'Please try again later or contact Digi-Key at https://www.digikey.com/en/help-support'
            self.show_error_message_then_exit(message=error_message, caption=error_caption)
        elif _data['state'] == 'CANNOT_LAUNCH_DEFAULT_BROWSER':
            error_caption = 'Cannot launch default browser'
            error_message = \
                'We\'ve created a Digi-Key list from the given part numbers, but could not open the list for you.\n' \
                'To view the list, you can open your preferred browser with this URL: {url}\n' \
                'Hint: You don\'t have to type the URL manually, ' \
                'Ctrl+C this message, then Ctrl+V elsewhere to get the message content, including the URL.' \
                ''.format(url=_data['url'])
            self.show_error_message_then_exit(message=error_message, caption=error_caption)
        else:
            self.wx_progress_text.SetLabel(_data['state'])
            self.wx_progress_gauge.SetValue(_data['gauge_int'])

    def post_bom_data(self, _bom=None):
        self.wx_push_btn.Disable()
        self.wx_progress_gauge.Show()
        self.wx_progress_text.Show()

        self.bom = {}
        for row_idx in range(self.wx_bom_lc.GetItemCount()):
            mpn = self.wx_bom_lc.GetItemText(row_idx, 1)
            qty = self.wx_bom_lc.GetItemText(row_idx, 2)
            cus_ref = self.wx_bom_lc.GetItemText(row_idx, 3)
            # note = self.wx_bom_lc.GetItemText(row_idx, note_col_idx)
            self.bom[mpn] = {
                'mpn': mpn,
                'qty': qty,
                'cusRef': cus_ref,
                'note': "",
            }

        PushThread(self, json_data=json_from_bom__with_pn_as_key(self.bom), list_name=self.list_name)

    def show_error_message_then_exit(self, message='Message content', caption='Caption'):
        self.wx_md = wx.MessageDialog(parent=None, message=message, caption=caption)
        self.wx_md.ShowModal()
        self.Destroy()

class DigiKeyMyListsPlugin(pcbnew.ActionPlugin):
    def __init__(self):
        self.name = 'Push to DigiKey myLists'
        self.category = 'Manufacturing'
        self.pcbnew_icon_support = hasattr(self, 'show_toolbar_button')
        self.show_toolbar_button = True
        self.icon_file_name = os.path.join(
            os.path.dirname(__file__), 'toolbar_icon.png')
        self.dark_icon_file_name = os.path.join(
            os.path.dirname(__file__), 'toolbar_icon.png')
        self.description = 'Push schematic components to Digi-Key myLists for easy and quick part ordering.'

    def Run(self):
        BOMFrame().Show()
