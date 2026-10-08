# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import json


def save_checked_ids(checked_ids, filepath):
    """
    checked_ids: dict {eid_key_int: ElementId}
    filepath:    str — full path, .json extension
    """
    data = {
        u'version': 1,
        u'type':    u'FilterMan_CheckedIds',
        u'ids':     [int(k) for k in checked_ids.iterkeys()]
    }
    with open(filepath, u'w') as f:
        json.dump(data, f, indent=2)


def load_checked_ids(filepath):
    """
    Returns list of int (eid keys).
    Raises ValueError if the file is not a FilterMan JSON.
    """
    with open(filepath, u'r') as f:
        data = json.load(f)
    if data.get(u'type') != u'FilterMan_CheckedIds':
        raise ValueError(u'Not a FilterMan selection file')
    return [int(i) for i in data.get(u'ids', [])]
