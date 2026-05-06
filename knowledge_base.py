"""
K12 销售知识库模块
提炼自《SPIN Selling》和《影响力》，提供离线规则引擎数据和 RAG 知识片段
"""

# ─── SPIN 问题模板库 ───────────────────────────────────────────────────────────

SPIN_TEMPLATES = {
    "situation_questions": {
        "grade_related": [
            "孩子目前几年级？",
            "在哪个学校就读？",
            "班级大概多少人？",
        ],
        "current_study": [
            "目前有在上课外辅导吗？",
            "每天做作业大概需要多长时间？",
            "家里谁主要辅导孩子学习？",
        ],
        "goals": [
            "对孩子升学有什么目标？",
            "希望孩子在哪些科目上重点提升？",
        ],
    },
    "problem_questions": {
        "academic": [
            "孩子在{subject}上有没有遇到什么困难？",
            "最近几次考试成绩是什么趋势？",
            "孩子对{subject}的学习兴趣怎么样？",
        ],
        "habits": [
            "孩子写作业时专注力如何？",
            "有没有拖延或磨蹭的情况？",
            "预习复习的习惯养成了吗？",
        ],
        "emotional": [
            "孩子有没有表现出厌学情绪？",
            "考试前会不会特别焦虑？",
            "对自己的学习有信心吗？",
        ],
    },
    "implication_questions": {
        "academic_chain": [
            "如果{subject}基础这学期没补上，下学期难度加大会不会更难追？",
            "成绩持续下滑的话，对升学选择会有什么影响？",
            "这个知识点是后面很多内容的基础，现在没掌握会不会造成连锁反应？",
        ],
        "habit_chain": [
            "拖延的习惯如果延续到初中/高中，学科增多压力更大会怎样？",
            "每天写作业到很晚，长期睡眠不足会不会影响白天的学习效率？",
            "没有自主学习能力，到了高年级完全靠家长盯还盯得过来吗？",
        ],
        "emotional_chain": [
            "如果厌学情绪不及时干预，会不会从一个科目扩散到所有科目？",
            "自信心受挫后，孩子会不会越来越不愿意尝试和提问？",
            "家庭因为学习的事经常起冲突，对亲子关系的影响大吗？",
        ],
        "cost_chain": [
            "如果现在不解决，等到问题更严重了再补，是不是要花更多的时间和费用？",
            "其他孩子都在进步，差距会不会越拉越大？",
        ],
    },
    "need_payoff_questions": {
        "academic": [
            "如果成绩能提升上来，对孩子的整体排名会有什么变化？",
            "基础打牢了之后，学新内容是不是会轻松很多？",
        ],
        "habits": [
            "如果孩子能养成主动学习的习惯，您觉得这会带来哪些变化？",
            "不用每天盯作业了，您的时间和精力是不是也能释放出来？",
        ],
        "confidence": [
            "孩子要是在学习上有了成就感，对他整个人的状态会有什么影响？",
            "如果学会了一套有效的学习方法，这对他以后的学习会有多大帮助？",
        ],
        "family": [
            "学习的事情理顺了，家庭氛围会不会好很多？",
            "看到孩子在进步，您的心情是不是也会轻松很多？",
        ],
    },
}

# ─── 影响力策略匹配表 ─────────────────────────────────────────────────────────

INFLUENCE_STRATEGIES = {
    "reciprocity": {
        "name": "互惠原理",
        "triggers": ["初次接触", "家长犹豫", "尚未建立信任"],
        "actions": [
            "提供免费学情诊断",
            "分享考试真题和学习资料",
            "主动告知升学政策变化",
            "课后额外反馈孩子表现",
        ],
        "timing": "建立信任阶段，在正式推荐课程之前",
    },
    "commitment_consistency": {
        "name": "承诺一致",
        "triggers": ["家长口头认可问题", "完成试听", "参加过讲座"],
        "actions": [
            "引导家长确认问题的重要性",
            "共同设定阶段目标",
            "回顾家长之前的积极表态",
            "从小承诺推进到大承诺",
        ],
        "timing": "家长已经有初步认可之后",
    },
    "social_proof": {
        "name": "社会认同",
        "triggers": ["家长在犹豫对比", "不确定效果", "价格异议"],
        "actions": [
            "分享同类型学员的提升案例",
            "展示班级报名情况",
            "引用家长好评和推荐",
            "展示升学数据",
        ],
        "timing": "家长有顾虑、需要参考时",
    },
    "liking": {
        "name": "喜好原则",
        "triggers": ["初次接触", "关系维护", "长期跟进"],
        "actions": [
            "寻找共同点（同龄孩子、同区域）",
            "真诚赞美家长对教育的重视",
            "展现合作姿态而非销售姿态",
            "保持适当的跟进频率",
        ],
        "timing": "贯穿始终，尤其是初期信任建立",
    },
    "authority": {
        "name": "权威原则",
        "triggers": ["需要建立专业信任", "家长质疑效果", "跟竞品对比"],
        "actions": [
            "展示教师资质和教学成果",
            "提供专业学情分析报告",
            "解读考试政策和大纲变化",
            "引用教育理论和研究数据",
        ],
        "timing": "建立信任阶段和方案推荐阶段",
    },
    "scarcity": {
        "name": "稀缺原则",
        "triggers": ["家长已认可但拖延", "促单阶段", "时间敏感"],
        "actions": [
            "提醒班级名额情况（须真实）",
            "提示优惠截止时间（须真实）",
            "强调学习窗口期的重要性",
            "说明优秀老师档期的稀缺性",
        ],
        "timing": "家长已经认可价值，需要推动行动时",
    },
}

# ─── 家长画像 × SPIN 侧重映射（来自知识库第一部分表格）────────────────────────

SPIN_PROFILE_FOCUS = {
    "焦虑型":    {"S": "少", "P": "轻触", "I": "慎用", "N": "多用",
                  "note": "已够焦虑，不要再放大问题，多聚焦积极解决方案"},
    "理性分析型": {"S": "适中", "P": "深入", "I": "多用", "N": "多用",
                  "note": "喜欢逻辑链条，用数据和因果关系打动"},
    "价格敏感型": {"S": "少", "P": "适中", "I": "多用", "N": "多用",
                  "note": "核心目标：让【不解决的成本】>【报班费用】"},
    "名校执念型": {"S": "少", "P": "适中", "I": "多用", "N": "适中",
                  "note": "把问题直接与目标学校录取要求挂钩"},
    "素质培养型": {"S": "适中", "P": "轻触", "I": "适中", "N": "多用",
                  "note": "聚焦学习能力与习惯养成的长期价值，不只谈分数"},
    "佛系型":    {"S": "适中", "P": "多用", "I": "多用", "N": "适中",
                  "note": "先用 P+I 唤醒问题意识，再推进需求转化"},
}

# ─── 家长画像 × 影响力原则优先级（来自知识库第三部分）────────────────────────

INFLUENCE_PROFILE_FOCUS = {
    "焦虑型":    ["authority", "social_proof"],
    "理性分析型": ["authority", "social_proof"],
    "价格敏感型": ["reciprocity", "scarcity", "social_proof"],
    "名校执念型": ["authority", "social_proof", "scarcity"],
    "素质培养型": ["liking", "authority"],
    "佛系型":    ["social_proof", "scarcity", "reciprocity"],
}


# ─── 公共 API ─────────────────────────────────────────────────────────────────

def get_kb_references(profile_type: str, data: dict) -> dict:
    """Return structured KB references for the 策略依据 panel."""
    spin_focus = SPIN_PROFILE_FOCUS.get(profile_type, SPIN_PROFILE_FOCUS["佛系型"])
    principle_keys = INFLUENCE_PROFILE_FOCUS.get(profile_type, ["liking", "authority"])

    principles = []
    for key in principle_keys:
        p = INFLUENCE_STRATEGIES[key]
        principles.append({
            "name": p["name"],
            "timing": p["timing"],
            "actions": p["actions"][:3],
        })

    concerns = data.get("concerns", [])
    subjects = data.get("subjects", [])
    subj = subjects[0] if subjects else "该科目"

    question_examples = []
    if profile_type in ("价格敏感型", "佛系型") or "补差" in concerns or "提分" in concerns:
        for q in SPIN_TEMPLATES["implication_questions"]["academic_chain"][:2]:
            question_examples.append(q.replace("{subject}", subj))
    if "冲刺名校" in concerns or "升学" in concerns:
        question_examples += SPIN_TEMPLATES["implication_questions"]["cost_chain"][:2]
    if profile_type in ("焦虑型", "素质培养型"):
        question_examples += SPIN_TEMPLATES["need_payoff_questions"]["confidence"][:2]
    if profile_type == "名校执念型":
        question_examples += SPIN_TEMPLATES["need_payoff_questions"]["academic"][:2]
    if not question_examples:
        question_examples += SPIN_TEMPLATES["need_payoff_questions"]["family"][:2]

    return {
        "profile_type": profile_type,
        "spin_focus": spin_focus,
        "influence_principles": principles,
        "spin_question_examples": question_examples[:4],
    }


def get_rag_snippets(profile_type: str, data: dict) -> str:
    """Generate compact KB text to inject into Gemini system prompt (RAG mode)."""
    concerns = data.get("concerns", [])
    subjects = data.get("subjects", [])
    subj = subjects[0] if subjects else "该科目"

    spin = SPIN_PROFILE_FOCUS.get(profile_type, SPIN_PROFILE_FOCUS["佛系型"])
    principle_keys = INFLUENCE_PROFILE_FOCUS.get(profile_type, ["authority", "liking"])

    lines = [
        "## 知识库参考（SPIN销售框架 + 影响力原则）",
        "",
        f"### {profile_type} 的 SPIN 提问策略",
        f"- S 情境问题：{spin['S']}（基本信息已收集，无需多问）",
        f"- P 问题发现：{spin['P']}",
        f"- I 暗示后果：{spin['I']}",
        f"- N 价值确认：{spin['N']}",
        f"- 策略要点：{spin['note']}",
        "",
        f"### {profile_type} 的核心影响力原则",
    ]

    for key in principle_keys[:3]:
        p = INFLUENCE_STRATEGIES[key]
        acts = " / ".join(p["actions"][:2])
        lines.append(f"- {p['name']}：{p['timing']}。具体：{acts}")

    lines.append("")

    impl_qs = []
    if "补差" in concerns or "提分" in concerns:
        for q in SPIN_TEMPLATES["implication_questions"]["academic_chain"][:2]:
            impl_qs.append(q.replace("{subject}", subj))
    if "冲刺名校" in concerns or "升学" in concerns:
        impl_qs += SPIN_TEMPLATES["implication_questions"]["cost_chain"][:2]
    if not impl_qs:
        impl_qs += SPIN_TEMPLATES["implication_questions"]["habit_chain"][:2]

    if impl_qs:
        lines.append("### 参考暗示问题（I类）")
        for q in impl_qs[:3]:
            lines.append(f"- {q}")
        lines.append("")

    if profile_type in ("焦虑型", "素质培养型"):
        nq = SPIN_TEMPLATES["need_payoff_questions"]["confidence"][:2]
    elif profile_type == "名校执念型":
        nq = SPIN_TEMPLATES["need_payoff_questions"]["academic"][:2]
    else:
        nq = SPIN_TEMPLATES["need_payoff_questions"]["habits"][:2]

    lines.append("### 参考价值确认问题（N类）")
    for q in nq:
        lines.append(f"- {q}")
    lines.append("")

    lines += [
        "### 跟进心理学节奏",
        "- 首次：互惠（给有价值内容，不谈报班）",
        "- 二次：社会认同（分享同类家庭案例）",
        "- 三次：权威+稀缺（专业报告+名额/时间提醒）",
        "",
        "### 应用要求",
        "请严格基于上述 SPIN 框架和影响力原则来分析，话术和跟进策略要",
        "体现具体的 SPIN 问题类型和影响力原则，给出有框架依据的实操建议。",
    ]

    return "\n".join(lines)
