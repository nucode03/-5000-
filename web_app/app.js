const DATA_URL = "./data/words.json";
const STORAGE_KEY = "oxford5000-mobile-completed-v1";

const state = {
  words: [],
  completed: new Set(),
  selectedLetter: "A",
  filter: "all",
  query: "",
  installPrompt: null,
};

const elements = {
  alphabetNav: document.querySelector("#alphabet-nav"),
  clearCompleted: document.querySelector("#clear-completed"),
  emptyState: document.querySelector("#empty-state"),
  filters: [...document.querySelectorAll(".filter")],
  installButton: document.querySelector("#install-button"),
  listCount: document.querySelector("#list-count"),
  listTitle: document.querySelector("#list-title"),
  progressBar: document.querySelector("#progress-bar"),
  progressText: document.querySelector("#progress-text"),
  searchInput: document.querySelector("#search-input"),
  template: document.querySelector("#word-card-template"),
  wordList: document.querySelector("#word-list"),
};

function loadCompleted() {
  try {
    const stored = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    if (Array.isArray(stored)) state.completed = new Set(stored);
  } catch {
    state.completed = new Set();
  }
}

function saveCompleted() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify([...state.completed]));
}

function availableLetters() {
  return [...new Set(state.words.map((word) => word.letter))].sort();
}

function visibleWords() {
  const query = state.query.trim().toLocaleLowerCase("en");
  return state.words.filter((word) => {
    const matchesLetter = query || word.letter === state.selectedLetter;
    const matchesQuery = !query || `${word.word} ${word.meaning} ${word.partOfSpeech}`.toLocaleLowerCase("en").includes(query);
    const isDone = state.completed.has(word.id);
    const matchesFilter = state.filter === "all" || (state.filter === "done" ? isDone : !isDone);
    return matchesLetter && matchesQuery && matchesFilter;
  });
}

function updateProgress() {
  const completed = state.completed.size;
  const percent = state.words.length ? (completed / state.words.length) * 100 : 0;
  elements.progressText.textContent = `${completed.toLocaleString("ko-KR")} / ${state.words.length.toLocaleString("ko-KR")}`;
  elements.progressBar.style.width = `${percent}%`;
  elements.clearCompleted.hidden = completed === 0;
}

function renderAlphabet() {
  const letters = availableLetters();
  elements.alphabetNav.replaceChildren(...letters.map((letter) => {
    const button = document.createElement("button");
    button.className = "letter-button";
    button.type = "button";
    button.textContent = letter;
    button.dataset.letter = letter;
    button.classList.toggle("active", state.selectedLetter === letter && !state.query);
    button.setAttribute("aria-pressed", String(state.selectedLetter === letter && !state.query));
    button.addEventListener("click", () => {
      state.selectedLetter = letter;
      state.query = "";
      elements.searchInput.value = "";
      render();
    });
    return button;
  }));
}

function toggleComplete(id) {
  if (state.completed.has(id)) state.completed.delete(id);
  else state.completed.add(id);
  saveCompleted();
  render();
}

function makeCard(word) {
  const fragment = elements.template.content.cloneNode(true);
  const card = fragment.querySelector(".word-card");
  const completion = fragment.querySelector(".complete-button");
  const done = state.completed.has(word.id);
  fragment.querySelector(".word").textContent = word.word;
  fragment.querySelector(".part-of-speech").textContent = word.partOfSpeech;
  fragment.querySelector(".meaning").textContent = word.meaning;
  completion.setAttribute("aria-pressed", String(done));
  completion.setAttribute("aria-label", `${word.word} ${done ? "완료 취소" : "학습 완료"}`);
  card.classList.toggle("is-complete", done);
  completion.addEventListener("click", () => toggleComplete(word.id));
  return fragment;
}

function renderList() {
  const words = visibleWords();
  const isSearching = Boolean(state.query.trim());
  elements.listTitle.textContent = isSearching ? "검색 결과" : state.selectedLetter;
  elements.listCount.textContent = `${words.length.toLocaleString("ko-KR")}개 단어`;
  elements.emptyState.hidden = words.length !== 0;
  elements.wordList.replaceChildren(...words.map(makeCard));
}

function renderFilters() {
  elements.filters.forEach((button) => {
    const active = button.dataset.filter === state.filter;
    button.classList.toggle("active", active);
    button.setAttribute("aria-pressed", String(active));
  });
}

function render() {
  renderAlphabet();
  renderFilters();
  updateProgress();
  renderList();
}

function setupControls() {
  elements.searchInput.addEventListener("input", (event) => {
    state.query = event.target.value;
    render();
  });
  elements.filters.forEach((button) => button.addEventListener("click", () => {
    state.filter = button.dataset.filter;
    render();
  }));
  elements.clearCompleted.addEventListener("click", () => {
    state.completed.clear();
    saveCompleted();
    render();
  });
  window.addEventListener("beforeinstallprompt", (event) => {
    event.preventDefault();
    state.installPrompt = event;
    elements.installButton.hidden = false;
  });
  elements.installButton.addEventListener("click", async () => {
    if (!state.installPrompt) return;
    state.installPrompt.prompt();
    await state.installPrompt.userChoice;
    state.installPrompt = null;
    elements.installButton.hidden = true;
  });
}

async function initialize() {
  try {
    const response = await fetch(DATA_URL);
    if (!response.ok) throw new Error(`Could not load word data (${response.status})`);
    state.words = await response.json();
    const validIds = new Set(state.words.map((word) => word.id));
    state.completed = new Set([...state.completed].filter((id) => validIds.has(id)));
    render();
  } catch (error) {
    elements.listTitle.textContent = "불러오기 실패";
    elements.listCount.textContent = "인터넷 연결 후 다시 열어 주세요.";
    console.error(error);
  }
}

loadCompleted();
setupControls();
initialize();

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("./service-worker.js"));
}
