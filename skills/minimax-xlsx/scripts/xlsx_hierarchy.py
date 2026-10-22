#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Validated static institution mapping; leaf facts roll upward once per ancestor."""
from xlsx_advanced import parse_month


class Hierarchy:
    def __init__(self,spec):
        required={'mapping_version','valid_from','valid_to','levels','report_levels','nodes'}
        if not isinstance(spec,dict) or set(spec)-required-{'coverage'} or not required<=set(spec):
            raise ValueError('Hierarchy requires mapping_version, valid_from/to, levels, report_levels and nodes')
        for key in ('mapping_version','valid_from','valid_to'):
            if not isinstance(spec[key],str) or not spec[key].strip():
                raise ValueError('Invalid hierarchy '+key)
        self.coverage=spec.get('coverage','complete')
        if self.coverage not in ('complete','observed'):
            raise ValueError('Hierarchy coverage must be complete or observed')
        self.start=parse_month(('text',spec['valid_from']),'hierarchy valid_from')
        self.end=parse_month(('text',spec['valid_to']),'hierarchy valid_to')
        if self.end<self.start:
            raise ValueError('Hierarchy validity interval is reversed')
        self.levels=spec['levels']; self.report_levels=spec['report_levels']
        if not isinstance(self.levels,list) or not 1<=len(self.levels)<=10 or any(not isinstance(level,str) or not level.strip() for level in self.levels) or len(set(self.levels))!=len(self.levels):
            raise ValueError('Hierarchy levels must be 1..10 unique ordered labels')
        if not isinstance(self.report_levels,list) or not self.report_levels or any(level not in self.levels for level in self.report_levels) or len(set(self.report_levels))!=len(self.report_levels):
            raise ValueError('Invalid/duplicate hierarchy report_levels')
        if not isinstance(spec['nodes'],list) or not 1<=len(spec['nodes'])<=5000:
            raise ValueError('Hierarchy needs 1..5000 institution nodes')
        self.nodes={}; self.children={}
        for node in spec['nodes']:
            if not isinstance(node,dict) or set(node)!={'id','parent','level','name'}:
                raise ValueError('Hierarchy nodes require id/parent/level/name')
            if any(not isinstance(node[key],str) or not node[key].strip() or len(node[key])>128 for key in ('id','level','name')):
                raise ValueError('Institution IDs/levels/names must be literal text')
            if node['parent'] is not None and (not isinstance(node['parent'],str) or not node['parent'].strip()):
                raise ValueError('Each institution has exactly one parent or null')
            if node['id'] in self.nodes:
                raise ValueError('Duplicate institution ID / multiple-parent declaration: '+node['id'])
            self.nodes[node['id']]=dict(node)
            self.children[node['id']]=[]
        for identifier,node in self.nodes.items():
            parent=node['parent']
            if parent is not None:
                if parent not in self.nodes:
                    raise ValueError('Unknown parent institution: '+parent)
                self.children[parent].append(identifier)
        # Parent walks are bounded by declared node count; cycles are never tolerated.
        for identifier in self.nodes:
            seen=set(); current=identifier
            while current is not None:
                if current in seen:
                    raise ValueError('Institution hierarchy cycle at '+current)
                seen.add(current); current=self.nodes[current]['parent']
        for identifier,node in self.nodes.items():
            if node['level'] not in self.levels:
                raise ValueError('Unknown institution level: '+node['level'])
            position=self.levels.index(node['level'])
            if node['parent'] is None and position!=0 or node['parent'] is not None and position!=self.levels.index(self.nodes[node['parent']]['level'])+1:
                raise ValueError('Hierarchy levels must follow adjacent declared order')
            if not self.children[identifier] and node['level']!=self.levels[-1]:
                raise ValueError('Leaf institutions must be at the deepest declared level')
        self.descendant_leaves={identifier:set() for identifier in self.nodes}
        for identifier in self.nodes:
            if not self.children[identifier]:
                current=identifier
                while current is not None:
                    self.descendant_leaves[current].add(identifier)
                    current=self.nodes[current]['parent']
        self.version=spec['mapping_version']

    def node(self,identifier,period,leaf=False):
        if not isinstance(identifier,str) or not isinstance(period,str):
            raise ValueError('Institution and mapping period must be literal text')
        if identifier not in self.nodes:
            raise ValueError('Unmapped institution: '+identifier)
        month=parse_month(('text',period),'institution mapping period')
        if not self.start<=month<=self.end:
            raise ValueError('Period outside static mapping validity; historical restatement is not automatic')
        if leaf and self.children[identifier]:
            raise ValueError('Parent/summary institution facts are refused; source must contain deepest leaf facts only: '+identifier)
        return self.nodes[identifier]

    def ancestors(self,identifier,period):
        self.node(identifier,period,leaf=True)
        output=[]; current=identifier
        while current is not None:
            node=self.nodes[current]
            if node['level'] in self.report_levels:
                output.append(node)
            current=node['parent']
        return output

    def summary(self):
        return {'mapping_version':self.version,'node_count':len(self.nodes),'levels':self.levels,
                'report_levels':self.report_levels,'historical_restatement':False,'coverage_policy':self.coverage,
                'source_policy':'deepest leaf facts only; each leaf contributes once to each selected ancestor'}
