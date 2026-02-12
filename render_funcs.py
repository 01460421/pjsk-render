#!/usr/bin/env python3
"""
render_funcs.py — 純渲染函數模組 (無 Discord 依賴)
本機 bot.py 和雲端 render_server.py 共用
"""
import json, os, math
from io import BytesIO
from datetime import datetime

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

from img_render import (
    render_table_image, render_info_card, render_message_box,
    render_help_image, render_line_chart, Theme, CJK_FONT, SERIF_FONT
)

# ========== 常數 ==========
ENERGY_MULTIPLIERS = {0:1,1:5,2:10,3:15,4:20,5:25,6:27,7:29,8:31,9:33,10:35}
TIME_SLOTS = [f"{h:02d}:00" for h in range(24)]
TRACKED_RANKS = [1,2,3,10,20,50,100]

# ========== 歌曲 DB ==========
SONG_DB = []

def load_song_db():
    global SONG_DB
    try:
        p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'song_db.json')
        if os.path.exists(p):
            with open(p, 'r', encoding='utf-8') as f:
                SONG_DB = json.load(f)
            print(f"[SongDB] Loaded {len(SONG_DB)} songs")
        else:
            print(f"[SongDB] song_db.json not found at {p}")
    except Exception as e:
        print(f"[SongDB] Error: {e}")

load_song_db()

# ========== EP 計算 ==========
def calc_song_score(diff_arr, live_type, power, skill_mag, s6):
    base = diff_arr[2]; fever = diff_arr[10]
    if live_type == 'multi':
        rate = base + fever * 0.5 + diff_arr[6] * skill_mag + diff_arr[7] * s6
    else:
        rate = base + fever * 0.5 + diff_arr[4] * skill_mag + diff_arr[5] * s6
    return int(rate * power * 4)

def calc_ep_value(live_type, score, event_rate, bonus, boost_rate, power=0, life=1000):
    if live_type == 'multi':
        score_part = int((score + 0.075 * power * 5) / 17000)
    else:
        life_part = min(1000, life) / 1000
        score_part = int((score * life_part) / 17000)
    return int((score_part + 123) * (event_rate / 100) * (bonus / 100 + 1)) * boost_rate

def find_push_plans(target_ep_gap, power, bonus, skill_mag=2.2, s6=2.2,
                    live_type='multi', life=1000, interval=50,
                    energy_options=None, top_n=10, border_speed=0):
    if energy_options is None:
        energy_options = [5, 7, 10]
    by_energy = {e: [] for e in energy_options}
    for song in SONG_DB:
        sid = song['id']; title = song['title']
        stime = song['time']; rate = song['rate']
        diffs = song.get('diffs', {})
        if stime <= 0: continue
        for dk, darr in diffs.items():
            if not darr or len(darr) < 11: continue
            score = calc_song_score(darr, live_type, power, skill_mag, s6)
            for energy in energy_options:
                boost = ENERGY_MULTIPLIERS.get(energy, 1)
                ep = calc_ep_value(live_type, score, rate, bonus, boost, power, life)
                if ep <= 0: continue
                cycle = stime + interval
                eph = ep * (3600 / cycle)
                plays = math.ceil(target_ep_gap / ep)
                time_min = plays * cycle / 60
                stamina = plays * energy
                adj_plays = plays; adj_time_min = time_min
                adj_stamina = stamina; catchable = True
                if border_speed > 0:
                    net_ep = ep - border_speed * (cycle / 3600)
                    if net_ep <= 0:
                        catchable = False; adj_plays = 99999
                        adj_time_min = 99999; adj_stamina = 99999
                    else:
                        adj_plays = math.ceil(target_ep_gap / net_ep)
                        adj_time_min = adj_plays * cycle / 60
                        adj_stamina = adj_plays * energy
                by_energy[energy].append({
                    'title': title, 'id': sid, 'diff': dk, 'lv': darr[0],
                    'time': stime, 'rate': rate,
                    'energy': energy, 'boost': boost,
                    'ep': ep, 'eph': round(eph),
                    'plays': plays, 'time_min': round(time_min, 1),
                    'stamina': stamina, 'score': score,
                    'adj_plays': adj_plays, 'adj_time_min': round(adj_time_min, 1),
                    'adj_stamina': adj_stamina, 'catchable': catchable
                })
    results = []
    for energy in energy_options:
        items = by_energy[energy]
        items.sort(key=lambda x: (-x['catchable'], x['adj_plays'], -x['eph']))
        seen = set(); count = 0
        for r in items:
            if not r['catchable']: continue
            key = (r['id'], r['diff'])
            if key not in seen:
                seen.add(key); results.append(r); count += 1
                if count >= top_n: break
    return results

# ========== 肘人方案圖片 ==========
def create_push_plan_image(plans, target_rank, target_score, current_ep, gap,
                           power, bonus, event_name="", border_info=None):
    bi = border_info or {}
    border_speed = bi.get('speed_1h') or bi.get('speed_3h') or bi.get('speed_24h') or 0
    border_speed_label = "1h" if bi.get('speed_1h') else ("3h" if bi.get('speed_3h') else "24h")
    grouped = {}
    for p in plans:
        e = p['energy']
        if e not in grouped: grouped[e] = []
        if len(grouped[e]) < 15: grouped[e].append(p)
    energies = sorted(grouped.keys())[:3]
    if not energies: return None
    n_energies = len(energies); n_rows_per = 5
    total_rows = n_energies * n_rows_per * 2
    has_border = border_speed > 0
    header_pt = 100 + (30 if has_border else 0)
    total_pt = header_pt + 56 + 25 + n_energies*2*38 + total_rows*15 + 50
    h = max(16, total_pt / 60)
    fig, ax = plt.subplots(figsize=(16, h))
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis('off')
    fig.set_facecolor(Theme.BG)
    ax.add_patch(FancyBboxPatch((0.015,0.015),0.97,0.97,boxstyle="round,pad=0,rounding_size=0.008",
        facecolor='none',edgecolor=Theme.GOLD_DARK,linewidth=2.0,transform=ax.transAxes,clip_on=False))
    ax.add_patch(FancyBboxPatch((0.028,0.028),0.944,0.944,boxstyle="round,pad=0,rounding_size=0.006",
        facecolor='none',edgecolor=Theme.GOLD_DARK,linewidth=0.5,alpha=0.3,transform=ax.transAxes,clip_on=False))
    ax.text(0.5,0.980,"Elbow Assistant",fontsize=22,fontweight='bold',color=Theme.INK,
            ha='center',va='top',transform=ax.transAxes,fontfamily=CJK_FONT)
    ax.plot([0.10,0.35],[0.965,0.965],color=Theme.GOLD_DARK,linewidth=1,alpha=0.4,transform=ax.transAxes,clip_on=False)
    ax.plot([0.65,0.90],[0.965,0.965],color=Theme.GOLD_DARK,linewidth=1,alpha=0.4,transform=ax.transAxes,clip_on=False)
    y = 0.952
    border_name = bi.get('name', '???')
    ax.text(0.05,y,f"No.{target_rank}",fontsize=18,fontweight='bold',color=Theme.GOLD,transform=ax.transAxes,fontfamily=CJK_FONT)
    ax.text(0.16,y,border_name,fontsize=12,color=Theme.INK,transform=ax.transAxes,fontfamily=CJK_FONT,va='bottom')
    ax.text(0.95,y,f"{target_score/10000:,.2f}W",fontsize=14,fontweight='bold',color=Theme.ROYAL_BLUE,ha='right',transform=ax.transAxes,fontfamily=CJK_FONT)
    y -= 0.028
    ax.text(0.05,y,"You",fontsize=11,color=Theme.INK_FADED,transform=ax.transAxes,fontfamily=CJK_FONT)
    ax.text(0.16,y,f"{current_ep/10000:,.2f}W",fontsize=11,color=Theme.INK,fontweight='bold',transform=ax.transAxes,fontfamily=CJK_FONT)
    ax.text(0.40,y,f"Gap: {gap/10000:,.2f}W ({gap:,} EP)",fontsize=10,color=Theme.HERALDIC_RED,transform=ax.transAxes,fontfamily=CJK_FONT)
    ax.text(0.95,y,f"Power {power:,}  |  Bonus {bonus}%",fontsize=9,color=Theme.INK_FADED,ha='right',transform=ax.transAxes,fontfamily=CJK_FONT)
    y -= 0.022
    if has_border:
        ax.text(0.05,y,"Border Speed",fontsize=10,color=Theme.INK_FADED,transform=ax.transAxes,fontfamily=CJK_FONT)
        ax.text(0.22,y,f"{border_speed/10000:,.4f}W/h ({border_speed_label})",fontsize=10,
                color=Theme.HERALDIC_RED,fontweight='bold',transform=ax.transAxes,fontfamily=CJK_FONT)
        for sp_key, sp_label in [('speed_1h','1h'),('speed_3h','3h'),('speed_24h','24h')]:
            spd = bi.get(sp_key, 0)
            if spd > 0 and sp_label != border_speed_label:
                ax.text(0.55+({'3h':0,'24h':0.18,'1h':0}.get(sp_label,0)),y,
                        f"{sp_label}: {spd/10000:,.4f}W/h",fontsize=8.5,color=Theme.INK_FADED,
                        transform=ax.transAxes,fontfamily=CJK_FONT)
        y -= 0.022; y -= 0.005
        _cx=0.5; _s=0.005
        ax.plot([0.05,_cx-0.015],[y,y],color=Theme.GOLD_DARK,linewidth=0.5,alpha=0.4,transform=ax.transAxes,clip_on=False)
        ax.fill([_cx,_cx+_s*1.5,_cx,_cx-_s*1.5,_cx],[y+_s,y,y-_s,y,y+_s],color=Theme.GOLD_DARK,alpha=0.4,transform=ax.transAxes)
        ax.plot([_cx+0.015,0.95],[y,y],color=Theme.GOLD_DARK,linewidth=0.5,alpha=0.4,transform=ax.transAxes,clip_on=False)
        y -= 0.020
    else:
        y -= 0.008
        ax.plot([0.05,0.95],[y,y],color=Theme.GOLD_DARK,linewidth=0.5,alpha=0.3,transform=ax.transAxes,clip_on=False)
        y -= 0.015
    diff_colors = {'E':Theme.FOREST_GREEN,'N':Theme.ROYAL_BLUE,'H':Theme.COPPER,
                   'X':Theme.HERALDIC_RED,'M':Theme.DEEP_PURPLE,'A':Theme.PINK}
    avail = y - 0.04; n_e = len(energies)
    fixed_cost = 2*0.024 + 0.020; energy_cost = n_e*2*0.036; gap_cost = n_e*2*0.006
    row_space = avail - fixed_cost - energy_cost - gap_cost
    row_h = min(0.021, max(0.012, row_space / (n_e * 2 * n_rows_per)))
    sections = [("【長效方案】EP效率優先", lambda x: -x['eph']),
                ("【短效方案】最快追上", lambda x: x.get('adj_plays', x['plays']))]
    for sec_idx, (sec_title, sort_key) in enumerate(sections):
        ax.text(0.5,y,sec_title,fontsize=13,fontweight='bold',color=Theme.GOLD_DARK,ha='center',transform=ax.transAxes,fontfamily=CJK_FONT)
        y -= 0.007
        ax.plot([0.15,0.85],[y,y],color=Theme.GOLD_DARK,linewidth=0.6,alpha=0.4,transform=ax.transAxes,clip_on=False)
        y -= 0.017
        for ei, energy in enumerate(energies):
            rows = sorted(grouped[energy], key=sort_key)[:n_rows_per]
            boost = ENERGY_MULTIPLIERS.get(energy, 1)
            ax.text(0.05,y,f"x{energy}火",fontsize=11,fontweight='bold',color=Theme.ROYAL_BLUE,transform=ax.transAxes,fontfamily=CJK_FONT)
            ax.text(0.14,y+0.002,f"(x{boost})",fontsize=8,color=Theme.INK_FADED,transform=ax.transAxes,fontfamily=CJK_FONT)
            top = rows[0] if rows else None
            if top:
                ap=top.get('adj_plays',top['plays']); at=top.get('adj_time_min',top['time_min'])
                t_s=f"{at/60:.1f}h" if at>=60 else f"{at:.0f}m"
                ax.text(0.95,y,f"Best: {top['title'][:10]} → {ap}場 / {t_s} / {top.get('adj_stamina',top['stamina'])}體",
                        fontsize=8,color=Theme.FOREST_GREEN,ha='right',transform=ax.transAxes,fontfamily=CJK_FONT,fontweight='bold')
            y -= 0.020
            for lb,cx in [("#",0.05),("Song",0.09),("Diff",0.42),("EP/Play",0.50),("EP/h",0.61),("Plays",0.72),("Time",0.81),("Stam",0.91)]:
                ax.text(cx,y,lb,fontsize=7.5,color=Theme.INK_FADED,transform=ax.transAxes,fontfamily=CJK_FONT,fontweight='bold')
            y -= 0.003
            ax.plot([0.05,0.97],[y,y],color=Theme.GOLD_DARK,linewidth=0.4,alpha=0.5,transform=ax.transAxes,clip_on=False)
            y -= row_h * 0.6
            for ri, r in enumerate(rows):
                if ri % 2 == 1:
                    ax.add_patch(FancyBboxPatch((0.04,y-row_h*0.35),0.93,row_h*0.95,
                        boxstyle="round,pad=0,rounding_size=0.003",facecolor=Theme.PARCHMENT_D,
                        edgecolor='none',alpha=0.3,transform=ax.transAxes,clip_on=False))
                fs = min(9, max(7, row_h * 500))
                dc = diff_colors.get(r['diff'], Theme.INK)
                ax.text(0.05,y,f"{ri+1}.",fontsize=fs,color=Theme.INK_FADED,transform=ax.transAxes,fontfamily=CJK_FONT)
                tn = r['title'][:12]+'..' if len(r['title'])>12 else r['title']
                ax.text(0.09,y,tn,fontsize=fs,color=Theme.INK,transform=ax.transAxes,fontfamily=CJK_FONT)
                ax.text(0.42,y,f"{r['diff']}{r['lv']}",fontsize=fs,color=dc,transform=ax.transAxes,fontfamily=CJK_FONT,fontweight='bold')
                ax.text(0.50,y,f"{r['ep']:,}",fontsize=fs,color=Theme.FOREST_GREEN,transform=ax.transAxes,fontfamily=CJK_FONT,fontweight='bold')
                ax.text(0.61,y,f"{r['eph']:,}",fontsize=fs,color=Theme.HERALDIC_RED,transform=ax.transAxes,fontfamily=CJK_FONT,fontweight='bold')
                ap=r.get('adj_plays',r['plays']); at=r.get('adj_time_min',r['time_min']); ast_=r.get('adj_stamina',r['stamina'])
                ax.text(0.72,y,str(ap),fontsize=fs,color=Theme.INK,transform=ax.transAxes,fontfamily=CJK_FONT)
                time_str=f"{at/60:.1f}h" if at>=60 else f"{at:.0f}m"
                ax.text(0.81,y,time_str,fontsize=fs,color=Theme.INK_LIGHT,transform=ax.transAxes,fontfamily=CJK_FONT)
                ax.text(0.91,y,str(ast_),fontsize=fs,color=Theme.DEEP_PURPLE,transform=ax.transAxes,fontfamily=CJK_FONT)
                y -= row_h
            y -= 0.006
        if sec_idx == 0:
            y -= 0.004; _cx=0.5; _s=0.004
            ax.plot([0.08,_cx-0.015],[y,y],color=Theme.GOLD_DARK,linewidth=0.5,alpha=0.3,transform=ax.transAxes,clip_on=False)
            ax.fill([_cx,_cx+_s*1.5,_cx,_cx-_s*1.5,_cx],[y+_s,y,y-_s,y,y+_s],color=Theme.GOLD_DARK,alpha=0.3,transform=ax.transAxes)
            ax.plot([_cx+0.015,0.92],[y,y],color=Theme.GOLD_DARK,linewidth=0.5,alpha=0.3,transform=ax.transAxes,clip_on=False)
            y -= 0.016
    y -= 0.005
    if has_border:
        ax.text(0.5,max(y,0.045),f"* 場次/時間/體力已含榜線追趕修正 (目標時速 +{border_speed/10000:,.4f}W/h)",
                fontsize=7.5,color=Theme.INK_FADED,ha='center',transform=ax.transAxes,fontfamily=CJK_FONT,style='italic')
        y -= 0.015
    if event_name:
        ax.text(0.5,max(y,0.030),event_name,fontsize=8,color=Theme.INK_FADED,ha='center',
                transform=ax.transAxes,fontfamily=CJK_FONT,style='italic')
    ax.text(0.5,0.012,datetime.now().strftime('%Y-%m-%d %H:%M:%S'),fontsize=7,color=Theme.TEXT_DIM,
            ha='center',transform=ax.transAxes,fontfamily=SERIF_FONT,alpha=0.5,style='italic')
    ax.text(0.96,0.030,'omega',fontsize=7,color=Theme.INK_FADED,ha='right',
            transform=ax.transAxes,fontfamily=SERIF_FONT,alpha=0.3,style='italic')
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf,format='png',facecolor=fig.get_facecolor(),bbox_inches='tight',dpi=140)
    plt.close(fig); buf.seek(0)
    return buf

# ========== 排名詳細圖 ==========
def create_ranking_detail_image(target, prev_p, next_p, event_name, history_data=None):
    rank=target.get('rank',0); sc=target.get('score',0)
    dp=f"{(prev_p.get('score',0)-sc)/10000:,.4f}W" if prev_p else "-"
    dn=f"{(sc-next_p.get('score',0))/10000:,.4f}W" if next_p else "-"
    last_sc=target.get('last_score',0); last_pt=f"{last_sc/10000:.4f}W" if last_sc else "-"
    h1=target.get('last_1h_stats') or {}; h3=target.get('last_3h_stats') or {}; h24=target.get('last_24h_stats') or {}
    speed_1h=f"{h1['speed']/10000:.4f}W/h" if h1.get('speed') else "-"
    speed_3h=f"{h3['speed']/10000:.4f}W/h" if h3.get('speed') else "-"
    speed_24h=f"{h24['speed']/10000:.4f}W/h" if h24.get('speed') else "-"
    avg_1h=f"{h1['average']/10000:.4f}W" if h1.get('average') else "-"
    count_1h=str(h1.get('count',0)) if h1.get('count') else "-"
    count_3h=str(h3.get('count',0)) if h3.get('count') else "-"
    count_24h=str(h24.get('count',0)) if h24.get('count') else "-"
    lpa=target.get('last_played_at','')
    if lpa:
        try:
            lpa_dt=datetime.fromisoformat(lpa.replace('Z','+00:00'))
            lpa_str=lpa_dt.strftime('%H:%M:%S')
        except: lpa_str=lpa[:19] if len(lpa)>=19 else lpa
    else: lpa_str="-"
    info_data=target.get('last_player_info') or {}
    card=info_data.get('card') or {}; profile=info_data.get('profile') or {}
    card_str=f"Lv{card['level']} MR{card['master_rank']}" if card.get('level') else "-"
    word=profile.get('word','-') or "-"
    has_hist=history_data and len(history_data)>=2
    if has_hist:
        fig,(ax1,ax2)=plt.subplots(1,2,figsize=(20,12),gridspec_kw={'width_ratios':[1,1.2]})
    else:
        fig,ax1=plt.subplots(figsize=(11,12))
    fig.set_facecolor(Theme.BG); ax1.axis('off')
    ax1.add_patch(FancyBboxPatch((0.02,0.02),0.96,0.96,boxstyle="round,pad=0,rounding_size=0.008",
        facecolor='none',edgecolor=Theme.GOLD_DARK,linewidth=2.0,transform=ax1.transAxes,clip_on=False))
    ax1.add_patch(FancyBboxPatch((0.035,0.035),0.93,0.93,boxstyle="round,pad=0,rounding_size=0.006",
        facecolor='none',edgecolor=Theme.GOLD_DARK,linewidth=0.6,alpha=0.35,transform=ax1.transAxes,clip_on=False))
    rank_color=Theme.GOLD if rank<=3 else Theme.INK
    ax1.text(0.5,0.955,f"No.{rank}",fontsize=34,fontweight='bold',color=rank_color,ha='center',va='top',transform=ax1.transAxes,fontfamily=CJK_FONT)
    ax1.plot([0.12,0.38],[0.925,0.925],color=Theme.GOLD_DARK,linewidth=1,alpha=0.5,transform=ax1.transAxes,clip_on=False)
    ax1.plot([0.62,0.88],[0.925,0.925],color=Theme.GOLD_DARK,linewidth=1,alpha=0.5,transform=ax1.transAxes,clip_on=False)
    ax1.text(0.5,0.91,target.get('name','-'),fontsize=16,color=Theme.INK,ha='center',va='top',transform=ax1.transAxes,fontfamily=CJK_FONT)
    y=0.85
    for lb,vl,vc in [("ID",str(target.get('userId','-')),Theme.INK_LIGHT),
                      ("Total Score",f"{sc/10000:,.4f}W",Theme.ROYAL_BLUE),
                      ("Last PT",last_pt,Theme.HERALDIC_RED),
                      ("Gap Above",dp,Theme.FOREST_GREEN),("Gap Below",dn,Theme.INK_LIGHT)]:
        ax1.text(0.08,y,lb,fontsize=12,color=Theme.INK_FADED,transform=ax1.transAxes,fontfamily=CJK_FONT)
        ax1.text(0.40,y,vl,fontsize=12.5,color=vc,transform=ax1.transAxes,fontweight='bold',fontfamily=CJK_FONT)
        y -= 0.050
    y -= 0.008; cx=0.5; s=0.007
    ax1.plot([0.08,cx-0.03],[y,y],color=Theme.GOLD_DARK,linewidth=0.6,alpha=0.5,transform=ax1.transAxes,clip_on=False)
    ax1.fill([cx,cx+s*1.5,cx,cx-s*1.5,cx],[y+s,y,y-s,y,y+s],color=Theme.GOLD_DARK,alpha=0.5,transform=ax1.transAxes)
    ax1.plot([cx+0.03,0.92],[y,y],color=Theme.GOLD_DARK,linewidth=0.6,alpha=0.5,transform=ax1.transAxes,clip_on=False)
    y -= 0.025
    ax1.text(0.5,y,"- Speed Chronicle -",fontsize=12,fontweight='bold',color=Theme.DEEP_PURPLE,ha='center',transform=ax1.transAxes,fontfamily=CJK_FONT,style='italic')
    y -= 0.045
    for period,spd,cnt in [("1h",speed_1h,count_1h),("3h",speed_3h,count_3h),("24h",speed_24h,count_24h)]:
        ax1.text(0.08,y,period,fontsize=11,color=Theme.INK_FADED,transform=ax1.transAxes,fontweight='bold',fontfamily=CJK_FONT)
        ax1.text(0.20,y,spd,fontsize=11.5,color=Theme.FOREST_GREEN,transform=ax1.transAxes,fontweight='bold',fontfamily=CJK_FONT)
        ax1.text(0.68,y,f"{cnt} games",fontsize=9.5,color=Theme.INK_FADED,transform=ax1.transAxes,fontfamily=CJK_FONT)
        y -= 0.044
    ax1.text(0.08,y,"1h Avg",fontsize=11,color=Theme.INK_FADED,transform=ax1.transAxes,fontfamily=CJK_FONT)
    ax1.text(0.20,y,avg_1h,fontsize=11.5,color=Theme.ROYAL_BLUE,transform=ax1.transAxes,fontweight='bold',fontfamily=CJK_FONT)
    y -= 0.05; y -= 0.005
    ax1.plot([0.08,0.92],[y,y],color=Theme.GOLD_DARK,linewidth=0.5,alpha=0.4,transform=ax1.transAxes,clip_on=False)
    y -= 0.025
    ax1.text(0.5,y,"- Adventurer Info -",fontsize=12,fontweight='bold',color=Theme.ROYAL_BLUE,ha='center',transform=ax1.transAxes,fontfamily=CJK_FONT,style='italic')
    y -= 0.045
    for lb,vl in [("Card",card_str),("Last Seen",lpa_str),("Motto",word[:20] if len(word)>20 else word)]:
        ax1.text(0.08,y,lb,fontsize=11,color=Theme.INK_FADED,transform=ax1.transAxes,fontfamily=CJK_FONT)
        ax1.text(0.30,y,vl,fontsize=11,color=Theme.INK,transform=ax1.transAxes,fontfamily=CJK_FONT)
        y -= 0.042
    ax1.text(0.5,0.04,event_name,fontsize=10,color=Theme.INK_FADED,ha='center',transform=ax1.transAxes,fontfamily=CJK_FONT,style='italic',
             bbox=dict(boxstyle='round,pad=0.4',facecolor=Theme.PARCHMENT_D,edgecolor=Theme.GOLD_DARK,linewidth=0.6,alpha=0.7))
    ax1.text(0.5,0.015,datetime.now().strftime('%Y-%m-%d %H:%M:%S'),fontsize=8,color=Theme.TEXT_DIM,ha='center',
             transform=ax1.transAxes,fontfamily=SERIF_FONT,alpha=0.5,style='italic')
    if has_hist:
        ax2.set_facecolor(Theme.PARCHMENT_L)
        times=[h['time'] for h in history_data]; scores=[h['score']/10000 for h in history_data]
        ax2.plot(range(len(scores)),scores,'-',color=Theme.ROYAL_BLUE,linewidth=2.5,solid_capstyle='round')
        ax2.plot(range(len(scores)),scores,'o',color=Theme.ROYAL_BLUE,markersize=7,markerfacecolor=Theme.PARCHMENT_L,markeredgewidth=2,zorder=5)
        ax2.fill_between(range(len(scores)),scores,alpha=0.08,color=Theme.ROYAL_BLUE)
        ax2.set_title(f"No.{rank} Score Chronicle",fontsize=15,fontweight='bold',color=Theme.INK,pad=15,fontfamily=CJK_FONT)
        ax2.set_ylabel('Score (W)',fontsize=11,color=Theme.INK_FADED,fontfamily=CJK_FONT)
        step=max(1,len(times)//10)
        ax2.set_xticks(range(0,len(times),step))
        lbls=[t.split(' ')[1][:5] if ' ' in t else t[-5:] for t in times[::step]]
        ax2.set_xticklabels(lbls,rotation=45,ha='right',fontsize=9,color=Theme.INK_FADED)
        ax2.tick_params(axis='y',colors=Theme.INK_FADED,labelsize=9)
        ax2.grid(True,alpha=0.2,linestyle='--',color=Theme.GOLD_DARK)
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,p: f'{x:,.1f}W' if x>=1 else f'{x:.1f}W'))
        if len(scores)>1:
            sr=max(scores)-min(scores); off=sr*0.1 if sr>0 else scores[-1]*0.05
            ax2.annotate(f'{scores[-1]:,.2f}W',xy=(len(scores)-1,scores[-1]),
                         xytext=(len(scores)-1.5,scores[-1]+off),fontsize=11,
                         color=Theme.HERALDIC_RED,fontweight='bold',fontfamily=CJK_FONT,
                         arrowprops=dict(arrowstyle='->',color=Theme.HERALDIC_RED,lw=1.5))
        ax2.spines['top'].set_visible(False); ax2.spines['right'].set_visible(False)
        ax2.spines['left'].set_color(Theme.BORDER); ax2.spines['bottom'].set_color(Theme.BORDER)
    plt.tight_layout()
    buf=BytesIO()
    plt.savefig(buf,format='png',facecolor=fig.get_facecolor(),bbox_inches='tight',dpi=140)
    plt.close(fig); buf.seek(0)
    return buf

# ========== 班表圖片 ==========
def create_schedule_image(dt, schedule, members=None, dpi=130, pjsk_center=""):
    """members: bot_data['members'] dict, 由呼叫端傳入"""
    members = members or {}
    headers = ["時間","車種","平均倍率","P1","P2(S6)","P3","P4","P5","外援","備註"]
    rows = []
    def get_bonus(p):
        if not p: return 0
        b = p.get('bonus', 0) or 0
        if b == 0 and p.get('user_id') and p['user_id'] in members:
            b = members[p['user_id']].get('bonus', 0)
        return b
    for hour in TIME_SLOTS:
        sh = schedule.get(hour, {})
        def fp(p):
            if not p: return ""
            b = get_bonus(p); name = p.get('name','')
            return f"{name}({b:.2f})" if b > 0 else name
        def fs6(p):
            if not p: return ""
            n=p.get("name",""); b=get_bonus(p)
            pw=p.get("s6_power") or p.get("power",0) or 0
            if pw==0 and p.get('user_id') and p['user_id'] in members:
                pw=members[p['user_id']].get('s6_power',0) or members[p['user_id']].get('power',0)
            if b > 0:
                return f"{n}({b:.2f}/{pw/10000:.2f}萬)" if pw>0 else f"{n}({b:.2f})"
            else:
                return f"{n}({pw/10000:.2f}萬)" if pw>0 else n
        rows.append([hour, sh.get("car_type","蝦"),
            f"{sh.get('avg_bonus',0):.2f}" if sh.get('avg_bonus') else "",
            "omega", fs6(sh.get("p2")), fp(sh.get("p3")),
            fp(sh.get("p4")), fp(sh.get("p5")),
            fp(sh.get("support")), sh.get("note","")])
    cw = [0.05,0.04,0.055,0.06,0.14,0.13,0.13,0.13,0.13,0.10]
    hc = [Theme.HEADER_BG]*4 + [Theme.DEEP_PURPLE] + [Theme.HEADER_BG]*5
    cc = {2: Theme.HERALDIC_RED, 3: Theme.COPPER, 4: Theme.DEEP_PURPLE}
    return render_table_image(title="私車班表", subtitle=dt, headers=headers, rows=rows,
        col_widths=cw, header_colors=hc, col_colors=cc,
        footer=f"P1: omega | P2: S6 | P3-P5: 推手 | {pjsk_center}", figsize=(20, 14), dpi=dpi)

# ========== 排名列表圖 ==========
def create_ranking_list_image(rankings, start, end, event_name, history_records=None):
    filtered = [p for p in rankings if start<=p.get('rank',0)<=end]
    if not filtered: return None
    now=datetime.now()
    headers=["排名","玩家名稱","總分","1h時速","上一局PT","場次(1h)","距前一名"]
    rows=[]
    for i,p in enumerate(filtered):
        rk=p.get('rank',0); sc=p.get('score',0)
        dp = f"{(filtered[i-1].get('score',0)-sc)/10000:.2f}W" if i>0 else "-"
        h1 = p.get('last_1h_stats') or {}
        speed_1h = f"{h1['speed']/10000:.2f}W/h" if h1.get('speed') else "-"
        last_sc = p.get('last_score', 0)
        last_pt = f"{last_sc/10000:.4f}W" if last_sc else "-"
        count_1h = str(h1.get('count', 0)) if h1.get('count') else "-"
        rows.append([f"#{rk}", p.get('name','-'), f"{sc/10000:,.4f}W",
                     speed_1h, last_pt, count_1h, dp])
    rh={i:'#E8D5A8' for i,p in enumerate(filtered) if p.get('rank',999)<=3}
    return render_table_image(title=f"活動排名 #{start}–{end}", subtitle=event_name,
        headers=headers, rows=rows, col_widths=[0.06,0.22,0.16,0.14,0.14,0.10,0.14],
        col_colors={0:Theme.RED,2:Theme.BLUE,3:Theme.GREEN,4:Theme.PURPLE}, row_highlights=rh,
        footer=f"更新: {now.strftime('%Y-%m-%d %H:%M:%S')} | 資料來源: hisekai.org",
        figsize=(18, max(6, 2.5+len(filtered)*0.55)))

# ========== 成員表圖 ==========
def create_member_table_image(members):
    if not members: return None
    sm = sorted(members.items(), key=lambda x: x[1].get('bonus',0), reverse=True)
    headers = ["#","名稱","倍率","綜合力","多開","二開","三開","S6倍率","S6綜合","備註"]
    rows = []
    for i,(uid,m) in enumerate(sm,1):
        pw=m.get('power',0) or 0; s6p=m.get('s6_power',0) or 0
        b2=m.get('bonus_2',0) or 0; b3=m.get('bonus_3',0) or 0; s6b=m.get('s6_bonus',0) or 0
        rows.append([str(i), m.get('name','-'), f"{m.get('bonus',0):.2f}",
            f"{pw/10000:.2f}萬" if pw>0 else "-", m.get('multi','單開'),
            f"{b2:.2f}" if b2>0 else "-", f"{b3:.2f}" if b3>0 else "-",
            f"{s6b:.2f}" if s6b>0 else "-", f"{s6p/10000:.2f}萬" if s6p>0 else "-",
            m.get('note','') or "-"])
    cc = {2:Theme.RED,5:Theme.RED,6:Theme.RED,7:Theme.PURPLE,8:Theme.PURPLE,3:Theme.BLUE}
    return render_table_image(title="成員資料", subtitle=f"共 {len(rows)} 人",
        headers=headers, rows=rows,
        col_widths=[0.03,0.14,0.06,0.08,0.05,0.06,0.06,0.06,0.08,0.08],
        col_colors=cc, figsize=(20, max(6, 2+len(rows)*0.4)))

# ========== 時數表圖 ==========
def create_hours_table_image(stats, today_str=None):
    sorted_s = sorted(stats.items(), key=lambda x: x[1]["total_hours"], reverse=True)
    headers = ["#","名稱","原推時數","S6時數","外援時數","合計"]
    rows = []
    for i,(uid,s) in enumerate(sorted_s,1):
        rows.append([str(i), s["name"], str(s["pusher_hours"]),
                     str(s["s6_hours"]), str(s.get("support_hours",0)), str(s["total_hours"])])
    if not today_str:
        today_str = datetime.now().strftime("%Y-%m-%d")
    cc = {2:Theme.GREEN, 3:Theme.PURPLE, 5:Theme.RED}
    return render_table_image(
        title="成員累計時數", subtitle=f"共 {len(rows)} 人 | 統計至 {today_str}",
        headers=headers, rows=rows,
        col_widths=[0.05,0.28,0.14,0.14,0.14,0.12],
        col_colors=cc, figsize=(14, max(5, 2+len(rows)*0.4)),
        footer="原推=P3-P5 | S6=P2 | 每個排班時段計1小時")

# ========== 排名走勢圖 ==========
def create_ranking_chart(records, rank=None, event_name=None):
    """records: ranking_history['records'], 由呼叫端傳入"""
    if len(records) < 2: return None
    if event_name:
        records = [r for r in records if r.get("event","") == event_name]
    if len(records) < 2: return None
    borders = [str(rank)] if rank else [str(r) for r in TRACKED_RANKS]
    cmap = {"1":Theme.HERALDIC_RED,"2":Theme.ORANGE,"3":Theme.GOLD,"10":Theme.FOREST_GREEN,
            "20":Theme.CYAN,"50":Theme.ROYAL_BLUE,"100":Theme.DEEP_PURPLE}
    datasets = []; longest_times = []
    for b in borders:
        ts = []; ss = []
        for rec in records:
            if b in rec.get("borders", {}):
                ts.append(rec["time"]); ss.append(rec["borders"][b]["score"]/10000)
        if len(ts) >= 2:
            datasets.append((f"T{b}", ss, cmap.get(b, Theme.TEXT_LIGHT)))
            if len(ts) > len(longest_times): longest_times = ts
    if not datasets: return None
    return render_line_chart(
        title=f"T{rank} 走勢" if rank else "榜線走勢", subtitle=event_name or "",
        x_labels=[t[5:13] for t in longest_times], datasets=datasets,
        y_label="分數 (萬)", y_formatter=lambda v,_: f"{v:,.0f}W" if v>=1 else f"{v:.1f}W")
