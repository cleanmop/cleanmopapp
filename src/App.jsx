import { supabase } from "./supabase";
import { useEffect, useState } from "react";
import * as XLSX from "xlsx";

function App() {

  const theme = {
  primary: "#0B5CFF",
  primaryDark: "#0047D6",
  background: "#F4F6FA",
  card: "#FFFFFF",
};

  const ui = {
  pageTitle: {
    fontSize: "28px",
    fontWeight: "700",
    marginBottom: "20px",
  },

  userCard: {
    background: "#fff",
    borderRadius: "16px",
    padding: "15px",
    marginBottom: "20px",
    boxShadow: "0 2px 8px rgba(0,0,0,0.08)",
  },

  menuGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "12px",
  },

  menuCard: {
    background: "#fff",
    border: "none",
    borderRadius: "16px",
    padding: "20px",
    textAlign: "center",
    fontSize: "16px",
    fontWeight: "600",
    boxShadow: "0 2px 8px rgba(0,0,0,0.08)",
    cursor: "pointer",
  },

  logoutBtn: {
    marginTop: "10px",
    width: "100%",
    padding: "10px",
    borderRadius: "10px",
    border: "none",
    background: "#ef4444",
    color: "#fff",
  },
};

  const managerList = [
    "없음", "이상현", "김동완", "김성민", "고재경", "김주덕",
    "조용화", "정철진", "송병훈", "임정은", "여원현", "김상민",
    "이정훈", "이중근", "박진배", "최종권", "이재현", "윤요찬",
    "이왕국", "하진영", "심승환", "김세한",
  ];

  const itemList = [
    "벽걸이", "스탠드", "1웨이", "4웨이", "홈멀티",
    "드럼", "드럼 19k ↓", "드럼2단", "통돌이", "통돌이 19k ↓",
    "아기사랑", "빌트인", "초음파드럼", "초음파통돌이",
    "냉장고양문형", "냉장고4도어", "냉장고일반", "냉장고빌트인",
    "김치냉장고 스탠드", "김치냉장고 4도어",
    "김치냉장고 뚜껑식", "김치냉장고 빌트인", "방문취소",
  ];

  const companyName = { S: "삼성", H: "HS", C: "클린맙", L: "LG" };

  const managerYear = {
    이상현: "24", 김성민: "24",
    김동완: "25", 조용화: "25", 정철진: "25", 송병훈: "25", 김주덕: "25",
    고재경: "26", 김상민: "26", 이정훈: "26", 여원현: "26",
    임정은: "26", 박진배: "26", 이중근: "26",
  };

  const vocPenalty = {
  minor: { label: "경미", point: 1 },
  normal: { label: "보통", point: 3 },
  serious: { label: "중대", point: 5 },
  critical: { label: "심각", point: 10 },
};
  const allowanceTable = {
    "25-S": { 벽걸이: 50000, 스탠드: 75000, "1웨이": 70000, "4웨이": 73000, 홈멀티: 125000, 드럼: 85000, "드럼 19k ↓": 85000, 드럼2단: 105000, 통돌이: 60000, "통돌이 19k ↓": 60000, 아기사랑: 50000, 빌트인: 85000, 초음파드럼: 55000, 초음파통돌이: 50000, 냉장고양문형: 55000, 냉장고4도어: 55000, 냉장고일반: 40000, 냉장고빌트인: 65000, "김치냉장고 스탠드": 40000, "김치냉장고 4도어": 50000, "김치냉장고 뚜껑식": 30000, "김치냉장고 빌트인": 65000, 방문취소: 10000 },
    "25-H": { 벽걸이: 40000, 스탠드: 65000, "1웨이": 50000, "4웨이": 75000, 홈멀티: 105000, 드럼: 80000, "드럼 19k ↓": 75000, 드럼2단: 0, 통돌이: 50000, "통돌이 19k ↓": 45000, 아기사랑: 30000, 빌트인: 0, 초음파드럼: 0, 초음파통돌이: 0, 냉장고양문형: 0, 냉장고4도어: 0, 냉장고일반: 0, 냉장고빌트인: 0, "김치냉장고 스탠드": 0, "김치냉장고 4도어": 0, "김치냉장고 뚜껑식": 0, "김치냉장고 빌트인": 0, 방문취소: 0 },
    "25-C": { 벽걸이: 50000, 스탠드: 85000, "1웨이": 60000, "4웨이": 90000, 홈멀티: 135000, 드럼: 100000, "드럼 19k ↓": 100000, 드럼2단: 150000, 통돌이: 70000, "통돌이 19k ↓": 70000, 아기사랑: 70000, 빌트인: 100000, 초음파드럼: 0, 초음파통돌이: 0, 냉장고양문형: 0, 냉장고4도어: 0, 냉장고일반: 0, 냉장고빌트인: 0, "김치냉장고 스탠드": 0, "김치냉장고 4도어": 0, "김치냉장고 뚜껑식": 0, "김치냉장고 빌트인": 0, 방문취소: 0 },
    "26-S": { 벽걸이: 45000, 스탠드: 70000, "1웨이": 65000, "4웨이": 70000, 홈멀티: 115000, 드럼: 80000, "드럼 19k ↓": 80000, 드럼2단: 100000, 통돌이: 55000, "통돌이 19k ↓": 55000, 아기사랑: 40000, 빌트인: 80000, 초음파드럼: 50000, 초음파통돌이: 40000, 냉장고양문형: 50000, 냉장고4도어: 50000, 냉장고일반: 40000, 냉장고빌트인: 60000, "김치냉장고 스탠드": 40000, "김치냉장고 4도어": 40000, "김치냉장고 뚜껑식": 30000, "김치냉장고 빌트인": 60000, 방문취소: 10000 },
    "26-H": { 벽걸이: 45000, 스탠드: 65000, "1웨이": 50000, "4웨이": 75000, 홈멀티: 110000, 드럼: 80000, "드럼 19k ↓": 75000, 드럼2단: 0, 통돌이: 50000, "통돌이 19k ↓": 45000, 아기사랑: 30000, 빌트인: 0, 초음파드럼: 0, 초음파통돌이: 0, 냉장고양문형: 0, 냉장고4도어: 0, 냉장고일반: 0, 냉장고빌트인: 0, "김치냉장고 스탠드": 0, "김치냉장고 4도어": 0, "김치냉장고 뚜껑식": 0, "김치냉장고 빌트인": 0, 방문취소: 0 },
    "26-C": { 벽걸이: 50000, 스탠드: 85000, "1웨이": 60000, "4웨이": 90000, 홈멀티: 135000, 드럼: 100000, "드럼 19k ↓": 100000, 드럼2단: 150000, 통돌이: 70000, "통돌이 19k ↓": 70000, 아기사랑: 70000, 빌트인: 100000, 초음파드럼: 0, 초음파통돌이: 0, 냉장고양문형: 0, 냉장고4도어: 0, 냉장고일반: 0, 냉장고빌트인: 0, "김치냉장고 스탠드": 0, "김치냉장고 4도어": 0, "김치냉장고 뚜껑식": 0, "김치냉장고 빌트인": 0, 방문취소: 0 },
    "24-S": { 벽걸이: 56500, 스탠드: 85800, "1웨이": 83000, "4웨이": 88000, 홈멀티: 142500, 드럼: 108000, "드럼 19k ↓": 108000, 드럼2단: 130000, 통돌이: 70000, "통돌이 19k ↓": 70000, 아기사랑: 60000, 빌트인: 110000, 초음파드럼: 64000, 초음파통돌이: 52000, 냉장고양문형: 70000, 냉장고4도어: 70000, 냉장고일반: 55000, 냉장고빌트인: 82000, "김치냉장고 스탠드": 60000, "김치냉장고 4도어": 70000, "김치냉장고 뚜껑식": 42000, "김치냉장고 빌트인": 82000, 방문취소: 10000 },
    "24-H": { 벽걸이: 52000, 스탠드: 80000, "1웨이": 61000, "4웨이": 85000, 홈멀티: 132000, 드럼: 90000, "드럼 19k ↓": 83000, 드럼2단: 0, 통돌이: 58000, "통돌이 19k ↓": 50000, 아기사랑: 38000, 빌트인: 0, 초음파드럼: 0, 초음파통돌이: 0, 냉장고양문형: 0, 냉장고4도어: 0, 냉장고일반: 0, 냉장고빌트인: 0, "김치냉장고 스탠드": 0, "김치냉장고 4도어": 0, "김치냉장고 뚜껑식": 0, "김치냉장고 빌트인": 0, 방문취소: 0 },
    "24-C": { 벽걸이: 60000, 스탠드: 120000, "1웨이": 90000, "4웨이": 120000, 홈멀티: 180000, 드럼: 160000, "드럼 19k ↓": 160000, 드럼2단: 200000, 통돌이: 100000, "통돌이 19k ↓": 100000, 아기사랑: 70000, 빌트인: 180000, 초음파드럼: 0, 초음파통돌이: 0, 냉장고양문형: 0, 냉장고4도어: 0, 냉장고일반: 0, 냉장고빌트인: 0, "김치냉장고 스탠드": 0, "김치냉장고 4도어": 0, "김치냉장고 뚜껑식": 0, "김치냉장고 빌트인": 0, 방문취소: 0 },
  };

  const [page, setPage] = useState("home");
  const [users, setUsers] = useState([]);
  const [currentUser, setCurrentUser] = useState(null);
  const [newUserName, setNewUserName] = useState("");
  const [newUserRole, setNewUserRole] = useState("manager");
  const [newUserYear, setNewUserYear] = useState("26");
  const [newPassword, setNewPassword] = useState("");
  const [loginName, setLoginName] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [selectedMonth, setSelectedMonth] = useState(new Date().toISOString().slice(0, 7));
  const [lockedMonths, setLockedMonths] = useState([]);
  const [selectedManager, setSelectedManager] = useState("전체");
  const [editingWorkId, setEditingWorkId] = useState(null);
  const [backPressed, setBackPressed] = useState(false);

  const [works, setWorks] = useState([]);

  const [workDate, setWorkDate] = useState("");
  const [company, setCompany] = useState("S");
  const [codeName, setCodeName] = useState("없음");
  const [workType, setWorkType] = useState("solo");
  const [teamCount, setTeamCount] = useState(2);
  const [visitManagers, setVisitManagers] = useState(Array(10).fill("없음"));
  const [items, setItems] = useState(Object.fromEntries(itemList.map((item) => [item, 0])));

  const [vocs, setVocs] = useState([]);

const [vocDate, setVocDate] = useState("");
const [vocCompany, setVocCompany] = useState("S");
const [vocManager, setVocManager] = useState("");
const [vocLevel, setVocLevel] = useState("minor");
const [vocMemo, setVocMemo] = useState("");

  useEffect(() => {
  const savedUser = localStorage.getItem("currentUser");

  if (savedUser) {
    const user = JSON.parse(savedUser);
    setCurrentUser(user);
    setSelectedManager(user.name);
  }

  loadWorks();
  loadVocs();
  loadUsers();
}, []);
  
  const monthlyWorks = works.filter((work) => work.workDate.startsWith(selectedMonth));

    const visibleWorks =
  currentUser?.role === "manager"
    ? monthlyWorks.filter((work) =>
        (work.visitManagers || []).includes(currentUser.name)
      )
    : monthlyWorks;

  const monthlyVocs = vocs.filter(
  (voc) => voc.date.startsWith(selectedMonth)
);

  const visibleVocs =
  currentUser?.role === "manager"
    ? monthlyVocs.filter((voc) => voc.manager === currentUser.name)
    : monthlyVocs;


  const changeItemCount = (itemName, value) => {
    setItems({ ...items, [itemName]: Number(value) });
  };

  const resetForm = () => {
    setWorkDate("");
    setCompany("S");
    setCodeName("없음");
    setWorkType("solo");
    setTeamCount(2);
    setItems(Object.fromEntries(itemList.map((item) => [item, 0])));
    setEditingWorkId(null);
  };

  const saveWork = async () => {
    if (!workDate) return alert("작업일을 선택하세요.");
    if (isMonthLocked(workDate.slice(0, 7))) {
  alert("마감된 월은 작업을 등록/수정할 수 없습니다.");
  return;
}
    if (!visitManagers[0] || visitManagers[0] === "없음") {
  return alert("방문매니저1을 선택하세요.");
}

    const selectedManagers =
  workType === "solo"
    ? [visitManagers[0]].filter((name) => name && name !== "없음")
    : visitManagers
        .slice(0, teamCount)
        .filter((name) => name && name !== "없음");

    if (workType === "solo" && selectedManagers.length !== 1) {
  return alert("혼자 작업은 방문매니저 1명만 선택하세요.");
}

if (workType === "team" && selectedManagers.length !== teamCount) {
  return alert(`동행인원 ${teamCount}명을 모두 선택하세요.`);
}

    const hasItem = Object.values(items).some((count) => count > 0);
    if (!hasItem) return alert("품목 수량을 1개 이상 입력하세요.");

    const finalCodeName =
  currentUser?.role === "manager" ? currentUser.name : codeName;


    const newWork = {
  id: editingWorkId || Date.now(),
  workDate,
  company,
  codeName: finalCodeName,
  workType,
  teamCount: selectedManagers.length,
  visitManagers: selectedManagers,
  items,
};
    const dbWork = {
  id: newWork.id,
  work_date: newWork.workDate,
  company: newWork.company,
  code_name: newWork.codeName,
  work_type: newWork.workType,
  team_count: newWork.teamCount,
  visit_managers: newWork.visitManagers,
  items: newWork.items,
};

if (editingWorkId) {
  const { error } = await supabase
    .from("works")
    .update(dbWork)
    .eq("id", editingWorkId);

  if (error) {
    console.error(error);
    alert(JSON.stringify(error));
    return;
  }

  await loadWorks();
} else {
  const { error } = await supabase
    .from("works")
    .insert([dbWork]);

  if (error) {
    console.error(error);
    alert(JSON.stringify(error));
    return;
  }

  await loadWorks();
}

    resetForm();
    alert(editingWorkId ? "수정되었습니다." : "저장되었습니다.");
  };

 const loadWorks = async () => {
  const { data, error } = await supabase
    .from("works")
    .select("*")
    .order("work_date", { ascending: false });

  if (error) {
    console.error(error);
    return;
  }
  
  useEffect(() => {
  const handlePopState = () => {
    if (page !== "home") {
      setPage("home");
      window.history.pushState(null, "", window.location.pathname);
      return;
    }

    if (!backPressed) {
      setBackPressed(true);

      alert("한 번 더 누르면 종료됩니다.");

      setTimeout(() => {
        setBackPressed(false);
      }, 2000);

      window.history.pushState(
        null,
        "",
        window.location.pathname
      );
    } else {
      window.history.back();
    }
  };

  window.history.pushState(
    null,
    "",
    window.location.pathname
  );

  window.addEventListener(
    "popstate",
    handlePopState
  );

  return () =>
    window.removeEventListener(
      "popstate",
      handlePopState
    );
}, [page, backPressed]);

  const converted = (data || []).map((work) => ({
    id: work.id,
    workDate: work.work_date,
    company: work.company,
    codeName: work.code_name,
    workType: work.work_type,
    teamCount: work.team_count,
    visitManagers: work.visit_managers,
    items: work.items,
  }));

  setWorks(converted);
};
  const deleteWork = async (id) => {
  const targetWork = works.find((work) => work.id === id);

  if (targetWork && isMonthLocked(targetWork.workDate.slice(0, 7))) {
    alert("마감된 월의 작업은 삭제할 수 없습니다.");
    return;
  }

  if (!confirm("이 작업을 삭제할까요?")) return;

  const { error } = await supabase
    .from("works")
    .delete()
    .eq("id", id);

  if (error) {
    console.error(error);
    alert(JSON.stringify(error));
    return;
  }

  await loadWorks();

  alert("삭제되었습니다.");
};

const loadVocs = async () => {
  const { data, error } = await supabase
    .from("vocs")
    .select("*")
    .order("voc_date", { ascending: false });

  if (error) {
    console.error(error);
    return;
  }

  const loadUsers = async () => {
  const { data, error } = await supabase
    .from("users")
    .select("*")
    .order("name");

  if (error) {
    console.error(error);
    return;
  }

  setUsers(data || []);
};

  const converted = (data || []).map((voc) => ({
    id: voc.id,
    date: voc.voc_date,
    company: voc.company,
    manager: voc.manager,
    level: voc.level,
    memo: voc.memo,
  }));

  setVocs(converted);
};

const loadUsers = async () => {
  const { data, error } = await supabase
    .from("users")
    .select("*")
    .order("name", { ascending: true });

  if (error) {
    console.error(error);
    return;
  }

  setUsers(data || []);
};

const loadMonthLocks = async () => {
  const { data, error } = await supabase
    .from("month_locks")
    .select("*");

  if (error) {
    console.error(error);
    return;
  }

  setLockedMonths(data || []);
};

const isMonthLocked = (month) => {
  return lockedMonths.some(
    (item) =>
      item.month === month &&
      item.locked === true
  );
};

const lockMonth = async () => {
  if (
    !confirm(
      `${selectedMonth}을(를) 마감할까요?`
    )
  )
    return;

  const { error } = await supabase
    .from("month_locks")
    .upsert([
      {
        month: selectedMonth,
        locked: true,
      },
    ]);

  if (error) {
    alert(JSON.stringify(error));
    return;
  }

  await loadMonthLocks();

  alert("월마감 완료");
};

const unlockMonth = async () => {
  if (
    !confirm(
      `${selectedMonth} 마감을 해제할까요?`
    )
  )
    return;

  const { error } = await supabase
    .from("month_locks")
    .delete()
    .eq("month", selectedMonth);

  if (error) {
    alert(JSON.stringify(error));
    return;
  }

  await loadMonthLocks();

  alert("마감 해제 완료");
};

  const startEditWork = (work) => {
    setEditingWorkId(work.id);
    setWorkDate(work.workDate);
    setCompany(work.company);
    setCodeName(work.codeName || "없음");
    setWorkType(work.workType || "solo");
    setTeamCount(work.teamCount || (work.visitManagers?.length || 1));

    const managers = work.visitManagers || [];
    const tempManagers = Array(10).fill("없음");

managers.forEach((name, index) => {
  tempManagers[index] = name;
});

setVisitManagers(tempManagers);
    setItems(work.items);
    setPage("workAdd");
  };

  const calculateSettlement = () => {
    const result = {};

    monthlyWorks.forEach((work) => {
      if (work.company === "L") return;

      const visitManagers = work.visitManagers || [];
      const peopleCount = visitManagers.length;
      if (peopleCount === 0) return;

      visitManagers.forEach((manager) => {
        const year = managerYear[manager];
        if (!year) return;

        const code = `${year}-${work.company}`;
        const allowance = allowanceTable[code];
        if (!allowance) return;

        if (!result[manager]) result[manager] = { total: 0, details: {} };

        Object.entries(work.items).forEach(([itemName, count]) => {
          if (count > 0) {
            const unitPay = allowance[itemName] || 0;
            const pay = Math.floor(
  (unitPay * count) / peopleCount
);

            if (!result[manager].details[itemName]) {
              result[manager].details[itemName] = { count: 0, amount: 0 };
            }

            result[manager].details[itemName].count += count / peopleCount;
            result[manager].details[itemName].amount += pay;
            result[manager].total += pay;
          }
        });
      });
    });

    return result;
  };

  const calculateManagerStats = () =>
    Object.entries(calculateSettlement())
      .map(([manager, data]) => ({ manager, total: data.total, details: data.details }))
      .sort((a, b) => b.total - a.total);

  const calculateMonthlyItems = (targetWorks = monthlyWorks) => {
  const result = {};

  targetWorks.forEach((work) => {
    Object.entries(work.items).forEach(([itemName, count]) => {
      if (count > 0) {
        result[itemName] = (result[itemName] || 0) + count;
      }
    });
  });

  return result;
};

  const getManagerWorks = (managerName) =>
    monthlyWorks.filter((work) => (work.visitManagers || []).includes(managerName));

  const calculateCompanyStats = () => {
    const result = {
      S: { workCount: 0, itemCount: 0 },
      H: { workCount: 0, itemCount: 0 },
      C: { workCount: 0, itemCount: 0 },
      L: { workCount: 0, itemCount: 0 },
    };

    monthlyWorks.forEach((work) => {
      result[work.company].workCount += 1;
      Object.values(work.items).forEach((count) => {
        if (count > 0) result[work.company].itemCount += count;
      });
    });

    const totalWorks = Object.values(result).reduce((sum, data) => sum + data.workCount, 0);
    const totalItems = Object.values(result).reduce((sum, data) => sum + data.itemCount, 0);

    return Object.entries(result).map(([company, data]) => ({
      company,
      workCount: data.workCount,
      itemCount: data.itemCount,
      workPercent: totalWorks > 0 ? ((data.workCount / totalWorks) * 100).toFixed(1) : "0.0",
      itemPercent: totalItems > 0 ? ((data.itemCount / totalItems) * 100).toFixed(1) : "0.0",
    }));
  };

   const saveVoc = async () => {
  if (!vocDate) return alert("등록일을 선택하세요.");
  if (isMonthLocked(vocDate.slice(0, 7))) {
  alert("마감된 월은 VOC를 등록할 수 없습니다.");
  return;
}
  if (!vocManager) return alert("매니저를 선택하세요.");
  if (!vocMemo) return alert("VOC 내용을 입력하세요.");

  const newVoc = {
    id: Date.now(),
    date: vocDate,
    company: vocCompany,
    manager: vocManager,
    level: vocLevel,
    memo: vocMemo,
  };

  const dbVoc = {
    id: newVoc.id,
    voc_date: newVoc.date,
    company: newVoc.company,
    manager: newVoc.manager,
    level: newVoc.level,
    memo: newVoc.memo,
  };

  const { error } = await supabase
    .from("vocs")
    .insert([dbVoc]);

  if (error) {
    console.error(error);
    alert(JSON.stringify(error));
    return;
  }

  await loadVocs();

  setVocDate("");
  setVocCompany("S");
  setVocManager("");
  setVocLevel("minor");
  setVocMemo("");

  alert("VOC가 등록되었습니다.");
};

const changePassword = async () => {
  if (!newPassword) {
    alert("새 비밀번호를 입력하세요.");
    return;
  }

  const { error } = await supabase
    .from("users")
    .update({ password: newPassword })
    .eq("id", currentUser.id);

  if (error) {
    console.error(error);
    alert("비밀번호 변경 실패");
    return;
  }

  alert("비밀번호가 변경되었습니다.");

const updatedUser = {
  ...currentUser,
  password: newPassword,
};

setCurrentUser(updatedUser);

setUsers(
  users.map((user) =>
    user.id === currentUser.id ? updatedUser : user
  )
);

setNewPassword("");
setPage("home");
};

const addUser = async () => {
  if (!newUserName) {
    alert("매니저명을 입력하세요.");
    return;
  }

  const { error } = await supabase
    .from("users")
    .insert([
      {
        name: newUserName,
        role: newUserRole,
        year: newUserYear,
        password: "0000",
      },
    ]);

  if (error) {
    alert(JSON.stringify(error));
    return;
  }

  await loadUsers();

  setNewUserName("");
  alert("매니저가 추가되었습니다.");
};

const resetUserPassword = async (id) => {
  if (!confirm("비밀번호를 0000으로 초기화할까요?")) return;

  const { error } = await supabase
    .from("users")
    .update({ password: "0000" })
    .eq("id", id);

  if (error) {
    alert(JSON.stringify(error));
    return;
  }

  await loadUsers();

  alert("초기화 완료");
};

const deleteUser = async (id) => {
  if (!confirm("매니저를 삭제할까요?")) return;

  const { error } = await supabase
    .from("users")
    .delete()
    .eq("id", id);

  if (error) {
    alert(JSON.stringify(error));
    return;
  }

  await loadUsers();

  alert("삭제 완료");
};

const deleteVoc = async (id) => {
  const targetVoc = vocs.find((voc) => voc.id === id);

  if (targetVoc && isMonthLocked(targetVoc.date.slice(0, 7))) {
    alert("마감된 월의 VOC는 삭제할 수 없습니다.");
    return;
  }

  if (!confirm("이 VOC를 삭제할까요?")) return;

  const { error } = await supabase
    .from("vocs")
    .delete()
    .eq("id", id);

  if (error) {
    console.error(error);
    alert(JSON.stringify(error));
    return;
  }

  await loadVocs();

  alert("삭제되었습니다.");
};

const calculateManagerScore = (manager) => {
  const totalPenalty = monthlyVocs
    .filter((voc) => voc.manager === manager)
    .reduce((sum, voc) => sum + (vocPenalty[voc.level]?.point || 0), 0);

  return Math.max(0, 100 - totalPenalty);
};

const getManagerVocs = (manager) => {
  return monthlyVocs.filter((voc) => voc.manager === manager);
};

const getManagerSummary = (managerName) => {
  const settlement = calculateSettlement();
  const managerData = settlement[managerName];

  const managerWorks = monthlyWorks.filter((work) =>
    (work.visitManagers || []).includes(managerName)
  );

  const managerVocs = monthlyVocs.filter(
    (voc) => voc.manager === managerName
  );

  return {
    workCount: managerWorks.length,
    totalPay: managerData?.total || 0,
    vocCount: managerVocs.length,
    score: calculateManagerScore(managerName),
  };
};

  const downloadSettlementExcel = () => {
    const rows = [];

    calculateManagerStats()
      .filter((data) => selectedManager === "전체" || data.manager === selectedManager)
      .forEach((data) => {
        Object.entries(data.details).forEach(([itemName, detail]) => {
          rows.push({
            정산월: selectedMonth,
            매니저: data.manager,
            품목: itemName,
            수량: detail.count,
            금액: Math.round(detail.amount),
          });
        });

        rows.push({
          정산월: selectedMonth,
          매니저: data.manager,
          품목: "합계",
          수량: "",
          금액: Math.round(data.total),
        });
      });

    const worksheet = XLSX.utils.json_to_sheet(rows);
    worksheet["!cols"] = [{ wch: 12 }, { wch: 12 }, { wch: 24 }, { wch: 10 }, { wch: 15 }];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "정산서");
    XLSX.writeFile(workbook, `클린맙_${selectedMonth}_정산서.xlsx`);
  };

  return (
    <div style={{ padding: "20px" }}>
      {!currentUser ? (
  <>
    <h2>로그인</h2>

    <select
      value={loginName}
      onChange={(e) => setLoginName(e.target.value)}
    >
      <option value="">사용자 선택</option>
      {users.map((user) => (
        <option key={user.id} value={user.name}>
          {user.name}
        </option>
      ))}
    </select>

    <br /><br />

    <input
      type="password"
      value={loginPassword}
      onChange={(e) => setLoginPassword(e.target.value)}
      placeholder="비밀번호"
    />

    <br /><br />

    <button
      onClick={() => {
        const selected = users.find(
          (user) =>
            user.name === loginName &&
            user.password === loginPassword
        );

        if (!selected) {
          alert("이름 또는 비밀번호가 맞지 않습니다.");
          return;
        }

        setCurrentUser(selected);
localStorage.setItem("currentUser", JSON.stringify(selected));

setSelectedManager(selected.name);
setLoginPassword("");
      }}
    >
      로그인
    </button>
  </>
) : null}

{page === "home" && currentUser && (
  <>
    <div
      style={{
        background: "linear-gradient(135deg, #0B5CFF, #0047D6)",
        color: "white",
        borderRadius: "0 0 34px 34px",
        padding: "32px 20px 42px",
        margin: "-20px -20px 20px",
        textAlign: "center",
      }}
    >
      <h1 style={{ margin: 0, fontSize: "26px", color:"white" }}>
        클린맙 업무관리 시스템
      </h1>
      <p style={{ margin: "8px 0 0", opacity: 0.9 }}>
        Cleanmop Management System
      </p>
    </div>

    <div
      style={{
        background: "white",
        borderRadius: "22px",
        padding: "18px",
        marginTop: "-45px",
        marginBottom: "18px",
        boxShadow: "0 4px 16px rgba(0,0,0,0.10)",
        display: "flex",
        alignItems: "center",
        gap: "14px",
      }}
    >
      <div
        style={{
          width: "58px",
          height: "58px",
          borderRadius: "50%",
          background: "#0B5CFF",
          color: "white",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontSize: "28px",
        }}
      >
        👤
      </div>

      <div>
        <div style={{ fontSize: "20px", fontWeight: "700" }}>
          {currentUser.name}님
        </div>
        <div style={{ color: "#666", marginTop: "4px" }}>
          {currentUser.role === "admin" ? "관리자" : "매니저"}
        </div>
      </div>
      <button
  onClick={() => {
    setCurrentUser(null);
    localStorage.removeItem("currentUser");
    setLoginName("");
    setLoginPassword("");
    setPage("home");
  }}
  style={{
    marginTop: "10px",
    border: "none",
    borderRadius: "10px",
    padding: "8px 12px",
    background: "#ef4444",
    color: "white",
  }}
>
  로그아웃
</button>
    </div>

    {currentUser.role === "manager" && (
      <div
        style={{
          background: "white",
          borderRadius: "22px",
          padding: "18px",
          marginBottom: "18px",
          boxShadow: "0 4px 16px rgba(0,0,0,0.08)",
        }}
      >
        <h3 style={{ marginTop: 0 }}>
          이번달 현황
        </h3>

        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            gap: "10px",
          }}
        >
          <div style={{ background: "#F4F7FF", borderRadius: "16px", padding: "14px", textAlign: "center" }}>
            <div style={{ fontSize: "22px", fontWeight: "800", color: "#0B5CFF" }}>
              {getManagerSummary(currentUser.name).workCount}
            </div>
            <div>작업 건수</div>
          </div>

          <div style={{ background: "#F4F7FF", borderRadius: "16px", padding: "14px", textAlign: "center" }}>
            <div style={{ fontSize: "22px", fontWeight: "800", color: "#0B5CFF" }}>
              {getManagerSummary(currentUser.name).vocCount}
            </div>
            <div>VOC 건수</div>
          </div>

          <div style={{ background: "#F4F7FF", borderRadius: "16px", padding: "14px", textAlign: "center" }}>
            <div style={{ fontSize: "18px", fontWeight: "800", color: "#0B5CFF" }}>
              {getManagerSummary(currentUser.name).totalPay.toLocaleString()}
            </div>
            <div>예상 정산금</div>
          </div>

          <div style={{ background: "#F4F7FF", borderRadius: "16px", padding: "14px", textAlign: "center" }}>
            <div style={{ fontSize: "22px", fontWeight: "800", color: "#0B5CFF" }}>
              {getManagerSummary(currentUser.name).score}
            </div>
            <div>매니저 점수</div>
          </div>
        </div>
      </div>
    )}

    <div
      style={{
        display: "grid",
        gridTemplateColumns: "1fr 1fr",
        gap: "14px",
      }}
    >
      {currentUser.role === "admin" && (
        <>
          <button style={ui.menuCard} onClick={() => setPage("workAdd")}>📋<br />작업 등록</button>
          <button style={ui.menuCard} onClick={() => setPage("workList")}>📂<br />작업 조회</button>
          <button style={ui.menuCard} onClick={() => setPage("settlement")}>💰<br />정산 조회</button>
          <button style={ui.menuCard} onClick={() => setPage("managerStats")}>📊<br />매니저 대시보드</button>
          <button style={ui.menuCard} onClick={() => setPage("vocAdd")}>⚠️<br />VOC 등록</button>
          <button style={ui.menuCard} onClick={() => setPage("vocList")}>📝<br />VOC 조회</button>
          <button style={ui.menuCard} onClick={() => setPage("companyStats")}>🏢<br />업체별 실적</button>
          <button style={ui.menuCard} onClick={() => setPage("userManage")}>👥<br />매니저 관리</button>
          <button style={ui.menuCard} onClick={() => setPage("monthLock")}>🔒<br />월마감</button>
          <button style={ui.menuCard} onClick={() => setPage("changePassword")}>🔑<br />비밀번호 변경</button>
        </>
      )}

      {currentUser.role === "manager" && (
        <>
          <button style={ui.menuCard} onClick={() => { setSelectedManager(currentUser.name); setPage("managerStats"); }}>📊<br />내 대시보드</button>
          <button style={ui.menuCard} onClick={() => setPage("workAdd")}>📋<br />작업 등록</button>
          <button style={ui.menuCard} onClick={() => setPage("workList")}>📂<br />내 작업 조회</button>
          <button style={ui.menuCard} onClick={() => setPage("vocList")}>⚠️<br />내 VOC 조회</button>
          <button style={ui.menuCard} onClick={() => setPage("changePassword")}>🔑<br />비밀번호 변경</button>
        </>
      )}
    </div>
  </>
)}

      {page === "workAdd" && (
        <>
          <div
  style={{
    position: "sticky",
    top: 0,
    backgroundColor: "white",
    padding: "10px",
    borderBottom: "1px solid #ddd",
    zIndex: 999,
  }}
>
  <button onClick={saveWork}>
    {editingWorkId ? "수정 저장" : "저장"}
  </button>

  {" "}

  <button onClick={() => setPage("home")}>
    메뉴
  </button>
</div>
          <h2>{editingWorkId ? "작업 수정" : "작업 등록"}</h2>

          <label>작업일</label><br />
          <input type="date" value={workDate} onChange={(e) => setWorkDate(e.target.value)} /><br /><br />

          <label>업체</label><br />
          <select value={company} onChange={(e) => setCompany(e.target.value)}>
            <option value="S">삼성</option>
            <option value="H">HS</option>
            <option value="C">클린맙</option>
            <option value="L">LG</option>
          </select><br /><br />

          <label>코드명</label><br />
          <select value={codeName} onChange={(e) => setCodeName(e.target.value)}>
            {managerList.map((name) => <option key={name} value={name}>{name}</option>)}
          </select><br /><br />

          <label>작업구분</label><br />
          <select
            value={workType}
            onChange={(e) => {
  setWorkType(e.target.value);
  setVisitManagers(Array(10).fill("없음"));
  setTeamCount(2);
}}
          >
            <option value="solo">혼자 작업</option>
            <option value="team">동행 작업</option>
          </select><br /><br />

          {workType === "team" && (
            <>
              <label>동행인원</label><br />
              <select value={teamCount} onChange={(e) => setTeamCount(Number(e.target.value))}>
  {Array.from({ length: 9 }, (_, i) => i + 2).map((num) => (
    <option key={num} value={num}>
      {num}명
    </option>
  ))}
</select><br /><br />
            </>
          )}

          {Array.from({
  length: workType === "solo" ? 1 : teamCount,
}).map((_, index) => (
  <div key={index}>
    <label>방문매니저 {index + 1}</label>
    <br />

    <select
      value={visitManagers[index]}
      onChange={(e) => {
        const updated = [...visitManagers];
        updated[index] = e.target.value;
        setVisitManagers(updated);
      }}
    >
      <option value="없음">선택</option>

      {managerList
        .filter((name) => name !== "없음")
        .map((name) => (
          <option key={name} value={name}>
            {name}
          </option>
        ))}
    </select>

    <br />
    <br />
  </div>
))}

          <hr />
          <h3>품목 수량</h3>

          {itemList.map((item) => (
            <div key={item} style={{ marginBottom: "12px" }}>
              <label>{item}</label><br />
              <input type="number" min="0" value={items[item]} onChange={(e) => changeItemCount(item, e.target.value)} />
            </div>
          ))}

          <button onClick={saveWork}>{editingWorkId ? "수정 저장" : "저장"}</button><br /><br />
          <button onClick={() => setPage("home")}>← 메뉴로 돌아가기</button>
        </>
      )}

      {page === "workList" && (
        <>
          <h2>작업 조회</h2>
          <input type="month" value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} />

          <div style={{ border: "1px solid #999", padding: "10px", margin: "15px 0", backgroundColor: "#f7f7f7" }}>
            <h3>{selectedMonth} 작업 요약</h3>
            <p>총 작업등록: {visibleWorks.length}건</p>
            {Object.entries(calculateMonthlyItems(visibleWorks)).map(([itemName, count]) => (
              <p key={itemName}>{itemName}: {count}대</p>
            ))}
          </div>

          {visibleWorks.map((work) => (
            <div key={work.id} style={{ border: "1px solid #ccc", padding: "10px", marginBottom: "10px" }}>
              <strong>{work.workDate}</strong>
              <p>업체: {companyName[work.company]}</p>
              <p>코드명: {work.codeName}</p>
              <p>작업구분: {work.workType === "team" ? `동행 작업 (${work.teamCount || work.visitManagers.length}명)` : "혼자 작업"}</p>
              <p>방문매니저: {(work.visitManagers || []).join(", ")}</p>
              <p>
                품목:
                {Object.entries(work.items)
                  .filter(([_, count]) => count > 0)
                  .map(([name, count]) => ` ${name} ${count}대`)
                  .join(", ")}
              </p>
              {!isMonthLocked(work.workDate.slice(0, 7)) && (
  <>
    <button onClick={() => startEditWork(work)}>수정</button>{" "}
    <button onClick={() => deleteWork(work.id)}>삭제</button>
  </>
)}
            </div>
          ))}

          <button onClick={() => setPage("home")}>← 메뉴로 돌아가기</button>
        </>
      )}

      {page === "settlement" && (
        <>
          <h2>정산 조회</h2>
          <input type="month" value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} />

          {Object.entries(calculateSettlement())
            .sort((a, b) => b[1].total - a[1].total)
            .map(([manager, data], index) => (
              <div key={manager} style={{ border: "1px solid #ccc", padding: "10px", marginBottom: "10px" }}>
                <h3>{index + 1}위. {manager}</h3>
                {Object.entries(data.details).map(([itemName, detail]) => (
                  <p key={itemName}><p>
  {itemName} : {detail.count.toFixed(1)}대
</p> : {detail.amount.toLocaleString()}원</p>
                ))}
                <strong>합계: {data.total.toLocaleString()}원</strong>
              </div>
            ))}

          <button onClick={() => setPage("home")}>← 메뉴로 돌아가기</button>
        </>
      )}

      {page === "managerStats" && (
        <>
          <h2>매니저별 대시보드</h2>
          <button onClick={downloadSettlementExcel}>엑셀 다운로드</button><br /><br />

          <label>조회월</label><br />
          <input type="month" value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} /><br /><br />

          {currentUser?.role === "admin" && (
  <>
    <label>매니저 선택</label><br />

    <select
      value={selectedManager}
      onChange={(e) => setSelectedManager(e.target.value)}
    >
      <option value="전체">전체</option>

      {managerList.map((name) => (
        <option key={name} value={name}>
          {name}
        </option>
      ))}
    </select>

    <br /><br />
  </>
)}
          {calculateManagerStats()
  .filter((data) => {
    if (currentUser?.role === "manager") {
      return data.manager === currentUser.name;
    }

    return selectedManager === "전체" || data.manager === selectedManager;
  })
  .map((data, index) => {
              
              return (
                <div key={data.manager} style={{ border: "1px solid #ccc", padding: "10px", marginBottom: "10px" }}>
                  <h3>{index + 1}위. {data.manager}</h3>
                  <p>VOC: {getManagerVocs(data.manager).length}건</p>
<p>매니저점수: {calculateManagerScore(data.manager)}점</p>

{getManagerVocs(data.manager).length > 0 && (
  <>
    <h4>VOC 상세내역</h4>

    {getManagerVocs(data.manager).map((voc) => (
      <div
        key={voc.id}
        style={{
          border: "1px solid #ddd",
          padding: "5px",
          marginTop: "5px",
          backgroundColor: "#fafafa",
        }}
      >
        <div>{voc.date}</div>
        <div>업체: {companyName[voc.company]}</div>
        <div>
          정도: {vocPenalty[voc.level]?.label} -
          {vocPenalty[voc.level]?.point}점
        </div>
        <div>내용: {voc.memo}</div>
      </div>
    ))}
  </>
)}

                  {Object.entries(data.details).map(([itemName, detail]) => (
                    <p key={itemName}>{itemName}: {detail.count}대 / {detail.amount.toLocaleString()}원</p>
                  ))}

                  <strong>총 정산: {data.total.toLocaleString()}원</strong>

                  <hr />
                  <h4>작업 내역</h4>
                  {getManagerWorks(data.manager).map((work) => (
                    <div key={work.id} style={{ border: "1px solid #ddd", padding: "5px", marginTop: "5px" }}>
                      <div>{work.workDate}</div>
                      <div>업체: {companyName[work.company]}</div>
                      <div>작업구분: {work.workType === "team" ? "동행 작업" : "혼자 작업"}</div>
                      <div>
                        품목:
                        {Object.entries(work.items)
                          .filter(([_, count]) => count > 0)
                          .map(([name, count]) => ` ${name} ${count}대`)
                          .join(", ")}
                      </div>
                    </div>
                  ))}
                </div>
              );
            })}

          <button onClick={() => setPage("home")}>← 메뉴로 돌아가기</button>
        </>
      )}

            {page === "companyStats" && (
        <>
          <h2>업체별 실적</h2>
          <input type="month" value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)} />

          {calculateCompanyStats().map((data) => (
            <div key={data.company} style={{ border: "1px solid #ccc", padding: "10px", margin: "10px 0" }}>
              <h3>{companyName[data.company]}</h3>
              <p>작업등록: {data.workCount}건</p>
              <p>작업등록 기준 점유율: {data.workPercent}%</p>
              <p>총 품목수량: {data.itemCount}대</p>
              <p>품목수량 기준 점유율: {data.itemPercent}%</p>
            </div>
          ))}

          <button onClick={() => setPage("home")}>← 메뉴로 돌아가기</button>
        </>
      )}
      {page === "vocAdd" && (
        <>
          <h2>VOC 등록</h2>

          <label>등록일</label><br />
          <input
            type="date"
            value={vocDate}
            onChange={(e) => setVocDate(e.target.value)}
          /><br /><br />

          <label>업체</label><br />
          <select value={vocCompany} onChange={(e) => setVocCompany(e.target.value)}>
            <option value="S">삼성</option>
            <option value="H">HS</option>
            <option value="C">클린맙</option>
            <option value="L">LG</option>
          </select><br /><br />

          <label>매니저명</label><br />
          <select value={vocManager} onChange={(e) => setVocManager(e.target.value)}>
            <option value="">선택</option>
            {managerList.filter((name) => name !== "없음").map((name) => (
              <option key={name} value={name}>{name}</option>
            ))}
          </select><br /><br />

          <label>VOC 정도</label><br />
          <select value={vocLevel} onChange={(e) => setVocLevel(e.target.value)}>
            <option value="minor">경미 -1점</option>
            <option value="normal">보통 -3점</option>
            <option value="serious">중대 -5점</option>
            <option value="critical">심각 -10점</option>
          </select><br /><br />

          <label>내용</label><br />
          <textarea
            value={vocMemo}
            onChange={(e) => setVocMemo(e.target.value)}
            placeholder="VOC 내용을 입력하세요"
            style={{ width: "300px", height: "100px" }}
          /><br /><br />

          <button onClick={saveVoc}>VOC 저장</button><br /><br />
          <button onClick={() => setPage("home")}>← 메뉴로 돌아가기</button>
        </>
      )}

      {page === "vocList" && (
        <>
          <h2>VOC 조회</h2>

          <label>조회월</label><br />
          <input
            type="month"
            value={selectedMonth}
            onChange={(e) => setSelectedMonth(e.target.value)}
          /><br /><br />

          {visibleVocs.length === 0 && <p>등록된 VOC가 없습니다.</p>}

          {visibleVocs.map((voc) => (
            <div key={voc.id} style={{ border: "1px solid #ccc", padding: "10px", margin: "10px 0" }}>
              <strong>{voc.date}</strong>
              <p>업체: {companyName[voc.company]}</p>
              <p>매니저: {voc.manager}</p>
              <p>정도: {vocPenalty[voc.level]?.label} -{vocPenalty[voc.level]?.point}점</p>
              <p>내용: {voc.memo}</p>
              {!isMonthLocked(voc.date.slice(0, 7)) && (
  <button onClick={() => deleteVoc(voc.id)}>삭제</button>
)}
            </div>
          ))}

          <button onClick={() => setPage("home")}>← 메뉴로 돌아가기</button>
        </>
      )}

          {page === "changePassword" && (
        <>
          <h2>비밀번호 변경</h2>

          <input
            type="password"
            value={newPassword}
            onChange={(e) => setNewPassword(e.target.value)}
            placeholder="새 비밀번호"
          />

          <br /><br />

          <button onClick={changePassword}>
            변경하기
          </button>

          <br /><br />

          <button onClick={() => setPage("home")}>
            ← 메뉴로 돌아가기
          </button>
        </>
      )} 

          {page === "userManage" && (
  <>
    <h2>매니저 관리</h2>

    <input
      value={newUserName}
      onChange={(e) => setNewUserName(e.target.value)}
      placeholder="매니저명"
    />

    <br /><br />

    <select
      value={newUserYear}
      onChange={(e) => setNewUserYear(e.target.value)}
    >
      <option value="24">24년차</option>
      <option value="25">25년차</option>
      <option value="26">26년차</option>
    </select>

    <br /><br />

    <button onClick={addUser}>
      매니저 추가
    </button>

    <hr />

    {users.map((user) => (
      <div key={user.id}>
        {user.name} ({user.role})

        <button
          onClick={() => resetUserPassword(user.id)}
        >
          비밀번호 초기화
        </button>

        <button
          onClick={() => deleteUser(user.id)}
        >
          삭제
        </button>

        <br /><br />
      </div>
    ))}

    <button onClick={() => setPage("home")}>
      ← 메뉴로 돌아가기
    </button>
  </>
)}

{page === "monthLock" && (
  <>
    <h2>월마감</h2>

    <input
      type="month"
      value={selectedMonth}
      onChange={(e) =>
        setSelectedMonth(e.target.value)
      }
    />

    <br /><br />

    <p>
      상태 :
      {isMonthLocked(selectedMonth)
        ? " 🔒 마감"
        : " 🔓 진행중"}
    </p>

    {!isMonthLocked(selectedMonth) ? (
  <button onClick={lockMonth}>
    월마감 실행
  </button>
) : (
  <button onClick={unlockMonth}>
    마감 해제
  </button>
)}

    <br /><br />

    <button onClick={() => setPage("home")}>
      ← 메뉴로 돌아가기
    </button>
  </>
)}
    </div>
  );
}

export default App;
