"""配色主题：蓝白（默认）/ 紫白。

用法：设环境变量 PPT_THEME=purple 再跑 dense_deck.py / paper_cover.py。
其它脚本（info_style / dense_deck / styled_deck / paper_cover）统一从这里取色，
不要在别处再写死色号。
"""
import os

THEMES = {
    'blue': dict(
        name='蓝白',
        navy='073DA5',
        layer=['073DA5', '1F56B8', '3E73CB', '6A93DA', '9DB8E8'],
        light='EAF0FC',
        pale='F8FAFE',
        arrow='BCD3F6',
        rule='D5DEEB',
        line='9CB3D3',
        ink='111722',
        red='DE1515',
        grey='6F7D90',
        head='92071C',
    ),
    'purple': dict(
        name='紫白',
        navy='5A2E8C',
        layer=['5A2E8C', '6E3FA3', '8559B8', 'A47FCC', 'C7B2E0'],
        light='F1EAF9',
        pale='FBF8FE',
        arrow='DCCBF0',
        rule='E2D8EF',
        line='C3AEDB',
        ink='14101A',
        red='C61A34',
        grey='6F6A78',
        head='7A1230',
    ),
    # 江南大学 VI 蓝（模板色 #0A418D）：主色直接取校色，浅蓝做底纹。PPT_THEME=jnu 生效。
    'jnu': dict(
        name='江南蓝',
        navy='0A418D',
        layer=['0A418D', '1D5AA8', '3D78C0', '6FA0D8', 'B4CCE8'],
        light='E8EFFA',
        pale='F7FAFE',
        arrow='B9CFEC',
        rule='D3DFEF',
        line='9BB4D4',
        ink='111722',
        red='C8161D',
        grey='6F7D90',
        head='92071C',
    ),
    # 参考稿（电解液竞赛稿）的深蓝—蓝绿双色：主色 #0B4A92，层阶向蓝绿 #1475A2 过渡，
    # 浅底取冷蓝灰。PPT_THEME=bank 生效。
    'bank': dict(
        name='金融蓝',
        navy='0B4A92',
        layer=['0B4A92', '10568F', '1475A2', '4A90B4', 'A9C9DA'],
        light='E9F2F9',
        pale='F8FBFD',
        arrow='BBD5E4',
        rule='D8E4EE',
        line='A2BED4',
        ink='111722',
        red='C8161D',
        grey='6F7D90',
        head='92071C',
    ),
    # 学术蓝（参考稿《学术蓝文献ppt》：深藏蓝 #183A6A + 蓝灰 #647B9B + 浅蓝灰底）：
    # 章节带、标题带、栏首统一走深藏蓝，浅底用 #E8EDF4，虚线框与细边框用 #647B9B。
    'academic': dict(
        name='学术蓝',
        navy='183A6A',
        layer=['183A6A', '254572', '33517B', '556E91', '8FA3BE'],
        light='E8EDF4',
        pale='F7F9FC',
        arrow='C3D3E8',
        rule='D6DEE9',
        line='647B9B',
        ink='1A1A1A',
        red='B3202C',
        grey='6E7A8A',
        head='183A6A',
    ),
    # 用户自带 VI（如中国农大“生命绿”）时用这档：主色深绿保证读性，浅绿做底纹，
    # 与校徽的亮绿（#00D200）属同色系。PPT_THEME=green 生效。
    'green': dict(
        name='生命绿',
        navy='127A3E',
        layer=['127A3E', '2A8F4F', '4CA76B', '86C79C', 'C2E4CD'],
        light='E8F7EC',
        pale='F7FDF8',
        arrow='BEE8C9',
        rule='D5EDDB',
        line='9FD3B0',
        ink='111722',
        red='C62828',
        grey='6F7D90',
        head='92071C',
    ),
    # 墨绿（医院 VI，取自院徽主色 #17493A）：标题带、色块与页眉统一走墨绿，
    # 章节带再压一档到 #0E3226。PPT_THEME=inkgreen 生效。
    'inkgreen': dict(
        name='墨绿',
        navy='17493A',
        layer=['12382C', '17493A', '2C6353', '5D8B79', 'A9C3B6'],
        light='E9F1EC',
        pale='F6FAF8',
        arrow='C0D6C9',
        rule='D7E2DA',
        line='9CB8A9',
        ink='151C18',
        red='B3202C',
        grey='6E7A72',
        head='0E3226',
    ),
    # 暖橙米色（协商民主 / 青少年心理健康汇报）：主色砖橙保证白字可读，浅米做底纹，
    # 强调用深砖红。PPT_THEME=warm 生效。
    'warm': dict(
        name='暖橙',
        navy='B24A1E',
        layer=['8C3112', 'B24A1E', 'C9663A', 'E09873', 'F2CBB7'],
        light='FBEDE1',
        pale='FEFAF6',
        arrow='EFC9AE',
        rule='EFD9C7',
        line='D9B79C',
        ink='2B2320',
        red='B3261E',
        grey='7A6A5F',
        head='8C3112',
    ),
}

THEME = os.environ.get('PPT_THEME', 'blue').strip().lower()
C = THEMES.get(THEME, THEMES['blue'])
