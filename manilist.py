# -*- coding: utf-8 -*-

import os
import sys
import re
import codecs
import argparse
import pandas as pd

SBS_DSP_CODE = r'8b6c9223-3e5f-421b-bd34-ca3926bd0cd3-PREFERRED'
MARUWA_DSP_CODE = r'c2bb1e2b-d818-4246-baab-72cb9e1d5ba8-PREFERRED'
FLEX_DSP_CODE = r''
route_dict = {}
info_dict = {}


class RouteInfo(object):
    name: str
    distance: float
    plantime: str
    routetime: str
    capacity: float
    servicetype: str
    dspcode: str
    count: int

    def __init__(self) -> None:
        self.name = ""
        self.distance = 0.00
        self.plantime = ""
        self.routetime = ""
        self.capacity = 0.00
        self.servicetype = ""
        self.dspcode = ""
        self.count = 0


def dump_route(route_name, useheader, showtimewindow, showaddress):
    if useheader is True:
        print(f'[{info_dict[route_name].name}]:'
              f' ItemCount:{info_dict[route_name].count}'
              f' ServiceType:{info_dict[route_name].servicetype}')
    _df: pd.DataFrame = route_dict[route_name]
    for _idx in range(len(_df)):
        print(_df['TRID'][_idx], end='')
        if showtimewindow:
            print(f',{str(_df["TimeWindow"][_idx])}', end='')
        if showaddress:
            print(f',"{str(_df["Address"][_idx])}"', end='')
        print()
    if useheader is True:
        print()
    return


def dump_schedule(route_name, useheader, showaddress):
    if useheader is True:
        print(f'[{info_dict[route_name].name}]:'
              f' ItemCount:{info_dict[route_name].count}'
              f' ServiceType:{info_dict[route_name].servicetype}'
              )
    _df: pd.DataFrame = route_dict[route_name]
    for _idx in range(len(_df)):
        if _df['TimeWindow'][_idx] != 'nan':
            print(_df['TRID'][_idx], end='')
            print(f',{str(_df["TimeWindow"][_idx])}', end='')
            if showaddress:
                print(f',"{str(_df["Address"][_idx]).decode("utf-8")}"', end='')
            print()
    if useheader is True:
        print()
    return


def dump_targets(routes, useheader, showtimewindow, showadress):
    for _route_name in routes:
        dump_route(_route_name, useheader, showtimewindow, showadress)
    return


def dump_schedule_targets(routes, useheader, showadress):
    for _route_name in routes:
        dump_schedule(_route_name, useheader, showadress)
    return


def expand_info(infostr: str) -> RouteInfo:
    _info: RouteInfo = RouteInfo()

    _ptn_info = re.compile(r'Route for ([A-Z]+[0-9]+)   '
                          r'Driving distance: ([0-9]+\.[0-9]+) km '
                          r'Route plan: ([0-9]+ hours [0-9]+ minutes) '
                          r'Route time: ([0-9]+ hours [0-9]+ minutes)   '
                          r'Total capacity: ([0-9]+\.[0-9]+) cu ft   '
                          r'Service type: (.+)   '
                          r'Preferred DSP: (.*)')
    _m_info = _ptn_info.fullmatch(infostr)
    _info.name = _m_info.group(1)
    _info.distance = float(_m_info.group(2))
    _info.plantime = _m_info.group(3)
    _info.routetime = _m_info.group(4)
    _info.capacity = float(_m_info.group(5))
    _info.servicetype = _m_info.group(6)
    _info.dspcode = _m_info.group(7)

    return _info


def get_dspcode(dspname: str):
    _dspcode: str = ''

    if dspname == 'sbs':
        _dspcode = SBS_DSP_CODE
    elif dspname == 'maruwa':
        _dspcode = MARUWA_DSP_CODE
    elif dspname == 'amflex':
        _dspcode = FLEX_DSP_CODE
    else:
        _dspcode = None

    return _dspcode


def load_file(filename):
    # Excelファイルを開く
    _db = pd.ExcelFile(filename)

    for _sheet in _db.sheet_names:
        _df = _db.parse(sheet_name=_sheet, skiprows=0, header=None)
        _infostr = _df[0].values[0] + ' ' + _df[4].values[0]
        _info: RouteInfo = expand_info(_infostr)
        info_dict[_info.name] = _info
        route_dict[_info.name] = _db.parse(
            _sheet,
            skiprows=3,
            skipfooter=1,
            usecols=[1, 4, 5],
            header=None,
            names=('TRID', 'TimeWindow', 'Address'))
        route_dict[_info.name]['TimeWindow'] = route_dict[_info.name]['TimeWindow'].astype(
            str)
        info_dict[_info.name].count = len(route_dict[_info.name])
    return

if __name__ == "__main__":
    #sys.stdout = codecs.getwriter('utf-8')(sys.stdout)
    #sys.stderr = codecs.getwriter('utf-8')(sys.stderr)
    _parser = argparse.ArgumentParser()

    _parser.add_argument('mode',
                        choices=['info', 'moreinfo', 'dump', 'dumpschedule'],
                        help='処理を選択します。(info=概要表示, dump=リスト表示)')
    _parser.add_argument('infile',
                        help='処理対象のmanifestファイルへのパス')
    _parser.add_argument('-d', '--dspname',
                        type=str,
                        choices=['sbs', 'maruwa', 'amflex', 'all'],
                        default='all',
                        help='処理対象のDSPを選択します。')
    _parser.add_argument('-t', '--targetroutes',
                        action='append',
                        help='処理対象のルート番号の一覧(ex: 1,2,5,10,... more)')
    _parser.add_argument('-n', '--noheader',
                        action='store_true')
    _parser.add_argument('-s', '--showtimewindow',
                        action='store_true')
    _parser.add_argument('-a', '--showaddress',
                        action='store_true')

    _args = _parser.parse_args()

    # 入力ファイル名
    if os.path.exists(_args.infile):
        load_file(_args.infile)
        _target_list = []
        if _args.dspname == 'all':
            # -tオプションはdspname='all'の時だけ検証
            if _args.targetroutes is None:
                # dspの指定が無く、かつ対象の指定が無いとき時
                # 全てのコースを設定する
                for _route_no in info_dict:
                    _target_list.append(_route_no)
            else:
                # -tにて指定されたリストを反映(存在しないルートは設定しない)
                for _targets in _args.targetroutes:
                    for _route_no in _targets.split(','):
                        if _route_no in info_dict:
                            _target_list.append(_route_no)
                        else:
                            print(f'指定されたルート"{_route_no}"は存在しません。スキップします。', file=sys.stderr)
        else:
            # 特定のDSPが選択された時の処理
            # DSP識別コードの取得
            _dspcode = get_dspcode(_args.dspname)
            _reject_list = []
            if _args.targetroutes is not None:
                for _rejects in _args.targetroutes:
                    for _route_no in _rejects.split(','):
                        if _route_no in info_dict:
                            _reject_list.append(_route_no)
                        else:
                            print(f'指定されたルート"{_route_no}"は存在しません。スキップします。', file=sys.stderr)

            for _route_no, _info in info_dict.items():
                # DSPが一致するルートのみ対象
                if _info.dspcode == _dspcode:
                    if not _route_no in _reject_list:
                        _target_list.append(_route_no)

        if _args.mode == 'info':
            print('Route, ItemCount')
            for _route_no in _target_list:
                print(",".join(
                    [
                        info_dict[_route_no].name,
                        str(info_dict[_route_no].count)
                    ])
                )
        elif _args.mode == 'moreinfo':
            print('Route, ItemCount, ServiceType, PlanTime, RouteTime')
            for _route_no in _target_list:

                print(','.join(
                    [
                        info_dict[_route_no].name,
                        str(info_dict[_route_no].count),
                        info_dict[_route_no].servicetype,
                        info_dict[_route_no].plantime,
                        info_dict[_route_no].routetime
                    ])
                )
        elif _args.mode == 'dump':
            dump_targets(_target_list, not _args.noheader,
                         _args.showtimewindow, _args.showaddress)
        elif _args.mode == 'dumpschedule':
            dump_schedule_targets(
                _target_list, not _args.noheader, _args.showaddress)
    else:
        print(f'ファイル"{_args.infile}"が見つかりません', file=sys.stderr)

    sys.exit()
