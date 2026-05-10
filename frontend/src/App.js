import React, { useState, useEffect } from 'react';
import { View, Text, StyleSheet, ScrollView, TouchableOpacity, Modal, TextInput, Switch } from 'react-native';
import { translations } from './translations';
import * as XLSX from 'xlsx-js-style';

const EXCEL_FONT_SIZE = 16;
const EXCEL_COL_NARROW_WCH = 10; // row number column only
const EXCEL_COL_WIDE_WCH = 48;

// RTL workbook view (Excel shows sheet right-to-left; col A on the right — matches Arabic layout)
const applyWorkbookRtl = (workbook) => {
  if (!workbook) return;
  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.Views = [{ RTL: true }];
};

/**
 * Column widths: narrow only for listed indexes (e.g. row-no column); wide for all others.
 * Cells: large font, vertical center, horizontal right + RTL reading order (numbers column centered).
 */
const finalizeExcelWorksheet = (worksheet, options = {}) => {
  const { narrowColumnIndexes = [] } = options;
  if (!worksheet || !worksheet['!ref']) return;
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  const numCols = range.e.c - range.s.c + 1;

  worksheet['!cols'] = [];
  for (let c = 0; c < numCols; c++) {
    const narrow = narrowColumnIndexes.includes(c);
    worksheet['!cols'][c] = { wch: narrow ? EXCEL_COL_NARROW_WCH : EXCEL_COL_WIDE_WCH };
  }

  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = worksheet[addr];
      if (!cell) continue;
      const narrow = narrowColumnIndexes.includes(C);
      cell.s = {
        font: { sz: EXCEL_FONT_SIZE },
        alignment: {
          horizontal: narrow ? 'center' : 'right',
          vertical: 'center',
          wrapText: true,
          readingOrder: 2
        }
      };
    }
  }
};

const API_URL = 'http://localhost:5000/api/persons';
const SEASONS_API = 'http://localhost:5000/api/seasons';

// Strip populated season subdocs for API save
const preparePersonBody = (person) => {
  if (!person) return {};
  const { activities, ...rest } = person;
  return {
    ...rest,
    activities: (activities || []).map((a) => ({
      description: a.description,
      date: a.date,
      status: a.status || 'pending',
      distributor: a.distributor != null ? String(a.distributor) : '',
      season: (() => {
        if (!a.season) return null;
        if (typeof a.season === 'object' && a.season._id) return a.season._id;
        return a.season;
      })()
    }))
  };
};

// Helper function to calculate age
const calculateAge = (birthday) => {
  if (!birthday) return null;
  const today = new Date();
  const birthDate = new Date(birthday);
  let age = today.getFullYear() - birthDate.getFullYear();
  const monthDiff = today.getMonth() - birthDate.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
    age--;
  }
  return age;
};

// Helper function to format date for input (YYYY-MM-DD)
const formatDateForInput = (dateString) => {
  if (!dateString) return '';
  const date = new Date(dateString);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

// Parse kid birthday into day, month, year for the three input boxes
const getBirthdayParts = (kid) => {
  if (!kid || !kid.birthday) return { day: '', month: '', year: '' };
  const date = new Date(kid.birthday);
  if (isNaN(date.getTime())) return { day: '', month: '', year: '' };
  return {
    day: String(date.getDate()),
    month: String(date.getMonth() + 1),
    year: String(date.getFullYear())
  };
};

// Build YYYY-MM-DD from day, month, year strings; return '' if invalid or incomplete
const buildBirthdayFromParts = (day, month, year) => {
  const d = day.trim();
  const m = month.trim();
  const y = year.trim();
  if (!d || !m || !y) return '';
  const dayNum = parseInt(d, 10);
  const monthNum = parseInt(m, 10);
  const yearNum = parseInt(y, 10);
  if (isNaN(dayNum) || isNaN(monthNum) || isNaN(yearNum)) return '';
  if (monthNum < 1 || monthNum > 12) return '';
  if (dayNum < 1 || dayNum > 31) return '';
  if (yearNum < 1900 || yearNum > 2100) return '';
  const monthStr = String(monthNum).padStart(2, '0');
  const dayStr = String(dayNum).padStart(2, '0');
  return `${yearNum}-${monthStr}-${dayStr}`;
};

const App = () => {
  const [language, setLanguage] = useState('ar'); // 'en' or 'ar'
  const t = translations[language];
  const isRTL = language === 'ar';
  const [currentView, setCurrentView] = useState('main'); // 'main' or 'openActivities'
  
  const [persons, setPersons] = useState([]);
  const [distributors, setDistributors] = useState(() => {
    // Load distributors from localStorage or initialize with empty array
    const saved = localStorage.getItem('distributors');
    return saved ? JSON.parse(saved) : [];
  });
  const [distributorModalVisible, setDistributorModalVisible] = useState(false);
  const [newDistributorName, setNewDistributorName] = useState('');
  const [newDistributorPhone, setNewDistributorPhone] = useState('');
  const [openDropdowns, setOpenDropdowns] = useState({}); // Track which dropdowns are open
  const [filteredPersons, setFilteredPersons] = useState([]);
  const [openActivitiesFilters, setOpenActivitiesFilters] = useState({
    distributor: '',
    season: '', // '' = all, '__NO_SEASON__', or season _id string
    activityStatus: 'open', // 'all', 'open', 'closed'
    activityName: '',
    dateFromDay: '',
    dateFromMonth: '',
    dateFromYear: '',
    dateToDay: '',
    dateToMonth: '',
    dateToYear: ''
  });
  const [showOpenActivitiesFilters, setShowOpenActivitiesFilters] = useState(false);
  const [paginatedPersons, setPaginatedPersons] = useState([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [itemsPerPage, setItemsPerPage] = useState(15);
  const [currentPage, setCurrentPage] = useState(1);
  const [showFilters, setShowFilters] = useState(false);
  const [filters, setFilters] = useState({
    minIncome: '',
    maxIncome: '',
    liveInRenta: null, // null = all, true = yes, false = no
    hasCar: null, // null = all, true = yes, false = no
    maritalStatus: null, // null = all, or 'Single', 'Married', 'Divorced', 'Widowed'
    kidsMinAge: '',
    kidsMaxAge: '',
    minKidsNumber: ''
  });
  const [modalVisible, setModalVisible] = useState(false);
  const [activitiesModalVisible, setActivitiesModalVisible] = useState(false);
  const [editingPerson, setEditingPerson] = useState(null);
  const [selectedPersonForActivities, setSelectedPersonForActivities] = useState(null);
  const [selectedPersonsForActivities, setSelectedPersonsForActivities] = useState([]); // Track selected persons for bulk activity addition
  const [bulkAddActivityModalVisible, setBulkAddActivityModalVisible] = useState(false);
  const [bulkAddActivityDescription, setBulkAddActivityDescription] = useState('');
  const [bulkAddActivityDistributor, setBulkAddActivityDistributor] = useState(''); // distributor name or ''
  const [bulkAddDistributorDropdownOpen, setBulkAddDistributorDropdownOpen] = useState(false);
  const [bulkAddActivitySeasonId, setBulkAddActivitySeasonId] = useState('');
  const [bulkAddSeasonDropdownOpen, setBulkAddSeasonDropdownOpen] = useState(false);
  const [seasons, setSeasons] = useState([]);
  const [seasonModalVisible, setSeasonModalVisible] = useState(false);
  const [newSeasonName, setNewSeasonName] = useState('');
  const [newActivitySeasonId, setNewActivitySeasonId] = useState('');
  const [activitySeasonDropdownOpen, setActivitySeasonDropdownOpen] = useState(false);
  const [expandedKids, setExpandedKids] = useState({}); // Track which persons have kids expanded
  const [newActivity, setNewActivity] = useState('');
  const [showFavoritesOnly, setShowFavoritesOnly] = useState(false);
  const [formData, setFormData] = useState({
    name: '',
    phone: '',
    kids: [],
    monthIncome: '',
    maritalStatus: 'Single',
    liveInRenta: false,
    hasCar: false,
    bankNumber: '',
    description: ''
  });

  useEffect(() => {
    fetchPersons();
    fetchSeasons();
  }, []);

  const fetchSeasons = async () => {
    try {
      const response = await fetch(SEASONS_API);
      const data = await response.json();
      setSeasons(Array.isArray(data) ? data : []);
    } catch (error) {
      console.error('Error fetching seasons:', error);
    }
  };

  // Save distributors to localStorage whenever they change
  useEffect(() => {
    localStorage.setItem('distributors', JSON.stringify(distributors));
  }, [distributors]);

  // Filter persons based on search query, filters, and favorites
  useEffect(() => {
    let filtered = [...persons];

    // Search filter (name or phone)
    if (searchQuery.trim()) {
      const query = searchQuery.toLowerCase().trim();
      filtered = filtered.filter(person => 
        person.name.toLowerCase().includes(query) ||
        person.phone.toLowerCase().includes(query)
      );
    }

    // Income filters
    if (filters.minIncome) {
      const minIncome = parseFloat(filters.minIncome);
      if (!isNaN(minIncome)) {
        filtered = filtered.filter(person => person.monthIncome >= minIncome);
      }
    }

    if (filters.maxIncome) {
      const maxIncome = parseFloat(filters.maxIncome);
      if (!isNaN(maxIncome)) {
        filtered = filtered.filter(person => person.monthIncome <= maxIncome);
      }
    }

    // Live in Renta filter
    if (filters.liveInRenta !== null) {
      filtered = filtered.filter(person => person.liveInRenta === filters.liveInRenta);
    }

    // Has Car filter
    if (filters.hasCar !== null) {
      filtered = filtered.filter(person => person.hasCar === filters.hasCar);
    }

    // Marital Status filter
    if (filters.maritalStatus !== null) {
      filtered = filtered.filter(person => person.maritalStatus === filters.maritalStatus);
    }

    // Kids age range filter
    const kidsMinAge = filters.kidsMinAge ? parseInt(filters.kidsMinAge) : null;
    const kidsMaxAge = filters.kidsMaxAge ? parseInt(filters.kidsMaxAge) : null;
    
    if (kidsMinAge !== null || kidsMaxAge !== null) {
      filtered = filtered.filter(person => {
        if (!person.kids || person.kids.length === 0) return false;
        return person.kids.some(kid => {
          const age = calculateAge(kid.birthday);
          if (age === null) return false;
          const meetsMin = kidsMinAge === null || age >= kidsMinAge;
          const meetsMax = kidsMaxAge === null || age <= kidsMaxAge;
          return meetsMin && meetsMax;
        });
      });
    }

    // Minimum number of kids filter
    if (filters.minKidsNumber) {
      const minKids = parseInt(filters.minKidsNumber);
      if (!isNaN(minKids)) {
        filtered = filtered.filter(person => (person.kidsNumber || 0) >= minKids);
      }
    }

    // Favorites filter
    if (showFavoritesOnly) {
      filtered = filtered.filter(person => person.favorite);
    }

    setFilteredPersons(filtered);
    setCurrentPage(1); // Reset to first page when filters change
  }, [searchQuery, filters, persons, showFavoritesOnly]);

  // Pagination logic
  useEffect(() => {
    if (itemsPerPage === 'All') {
      setPaginatedPersons(filteredPersons);
    } else {
      const itemsPerPageNum = parseInt(itemsPerPage);
      const startIndex = (currentPage - 1) * itemsPerPageNum;
      const endIndex = startIndex + itemsPerPageNum;
      setPaginatedPersons(filteredPersons.slice(startIndex, endIndex));
    }
  }, [filteredPersons, itemsPerPage, currentPage]);

  const fetchPersons = async () => {
    try {
      const response = await fetch(API_URL);
      const data = await response.json();
      setPersons(data);
      setFilteredPersons(data);
    } catch (error) {
      console.error('Error fetching persons:', error);
    }
  };

  // Distributor management functions
  const addDistributor = () => {
    if (!newDistributorName.trim() || !newDistributorPhone.trim()) {
      alert(t.distributorNamePhoneRequired);
      return;
    }
    const newDistributor = {
      id: Date.now().toString(),
      name: newDistributorName.trim(),
      phone: newDistributorPhone.trim()
    };
    setDistributors([...distributors, newDistributor]);
    setNewDistributorName('');
    setNewDistributorPhone('');
    setDistributorModalVisible(false);
  };

  const removeDistributor = (distributorId) => {
    if (window.confirm(t.confirmRemoveDistributor)) {
      setDistributors(distributors.filter(d => d.id !== distributorId));
    }
  };

  // Get all people with their activities (one row per activity)
  const getAllPeopleWithActivities = () => {
    const peopleActivities = [];
    persons.forEach(person => {
      if (person.activities && person.activities.length > 0) {
        person.activities.forEach((activity, activityIndex) => {
          const seasonObj = activity.season && typeof activity.season === 'object' ? activity.season : null;
          const seasonId = seasonObj?._id || activity.season || null;
          const seasonName = seasonObj?.name || null;
          peopleActivities.push({
            ...activity,
            personId: person._id,
            personName: person.name,
            personPhone: person.phone || '',
            personKidsNumber: person.kidsNumber || 0,
            activityIndex: activityIndex,
            activityStatus: activity.status || 'pending',
            seasonId,
            seasonName
          });
        });
      } else {
        // Include people with no activities
        peopleActivities.push({
          personId: person._id,
          personName: person.name,
          personPhone: person.phone || '',
          personKidsNumber: person.kidsNumber || 0,
          description: '',
          distributor: '',
          activityStatus: 'none',
          date: person.createdAt || new Date().toISOString(),
          seasonId: null,
          seasonName: null
        });
      }
    });
    // Sort by date (oldest first)
    return peopleActivities.sort((a, b) => new Date(a.date) - new Date(b.date));
  };

  // Filter activities based on filters
  const getFilteredActivities = () => {
    let activities = getAllPeopleWithActivities();

    // Filter by distributor
    if (openActivitiesFilters.distributor) {
      if (openActivitiesFilters.distributor === '__NO_DISTRIBUTOR__') {
        // Filter for activities without distributor
        activities = activities.filter(activity => 
          !activity.distributor || activity.distributor === ''
        );
      } else {
        // Filter for specific distributor
        activities = activities.filter(activity => 
          activity.distributor === openActivitiesFilters.distributor
        );
      }
    }

    // Filter by activity status
    if (openActivitiesFilters.activityStatus === 'open') {
      activities = activities.filter(activity => 
        activity.activityStatus === 'pending' || !activity.activityStatus
      );
    } else if (openActivitiesFilters.activityStatus === 'closed') {
      activities = activities.filter(activity => 
        activity.activityStatus === 'completed'
      );
    }
    // 'all' shows everything

    // Filter by activity name/description
    if (openActivitiesFilters.activityName.trim()) {
      const searchTerm = openActivitiesFilters.activityName.toLowerCase().trim();
      activities = activities.filter(activity => 
        activity.description && activity.description.toLowerCase().includes(searchTerm)
      );
    }

    // Filter by season
    if (openActivitiesFilters.season) {
      if (openActivitiesFilters.season === '__NO_SEASON__') {
        activities = activities.filter(activity =>
          activity.description && !activity.seasonId
        );
      } else {
        const sid = openActivitiesFilters.season;
        activities = activities.filter(activity =>
          activity.seasonId && String(activity.seasonId) === String(sid)
        );
      }
    }

    // Timeline: activity date between from/to (day/month/year), inclusive — same rules as kid birthday fields
    const df = buildBirthdayFromParts(
      openActivitiesFilters.dateFromDay || '',
      openActivitiesFilters.dateFromMonth || '',
      openActivitiesFilters.dateFromYear || ''
    );
    const dt = buildBirthdayFromParts(
      openActivitiesFilters.dateToDay || '',
      openActivitiesFilters.dateToMonth || '',
      openActivitiesFilters.dateToYear || ''
    );
    if (df || dt) {
      let fromTs = null;
      let toTs = null;
      if (df) {
        const [y, m, d] = df.split('-').map(Number);
        fromTs = new Date(y, m - 1, d, 0, 0, 0, 0).getTime();
      }
      if (dt) {
        const [y, m, d] = dt.split('-').map(Number);
        toTs = new Date(y, m - 1, d, 23, 59, 59, 999).getTime();
      }
      activities = activities.filter((activity) => {
        const t = new Date(activity.date).getTime();
        if (Number.isNaN(t)) return false;
        if (fromTs !== null && t < fromTs) return false;
        if (toTs !== null && t > toTs) return false;
        return true;
      });
    }

    return activities;
  };

  const groupActivitiesBySeason = (activities) => {
    const groups = [];
    const indexByKey = {};
    activities.forEach((activity) => {
      const key = activity.seasonId ? String(activity.seasonId) : '__none__';
      if (indexByKey[key] === undefined) {
        indexByKey[key] = groups.length;
        groups.push({
          key,
          title: activity.seasonName || t.uncategorizedSeason,
          activities: []
        });
      }
      groups[indexByKey[key]].activities.push(activity);
    });
    return groups;
  };

  // Export filtered activities to Excel (No., Name, Kids number, Activity, Distributor only)
  const exportToExcel = () => {
    const filteredActivities = getFilteredActivities();

    const excelData = filteredActivities.map((activity, index) => ({
      [t.rowNo]: index + 1,
      [t.name]: activity.personName,
      [t.kids]: activity.personKidsNumber,
      [t.season]: activity.seasonName || t.uncategorizedSeason,
      [t.activity]: activity.description || '-',
      [t.distributor]: activity.distributor || '-'
    }));

    const worksheet = XLSX.utils.json_to_sheet(excelData);
    finalizeExcelWorksheet(worksheet, { narrowColumnIndexes: [0] });
    const workbook = XLSX.utils.book_new();
    applyWorkbookRtl(workbook);
    XLSX.utils.book_append_sheet(workbook, worksheet, t.openActivities);

    const fileName = `Open_Activities_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  };

  // Selected rows on Open Activities page: same columns as exportToExcel, only checked people, still respecting filters
  const exportSelectedOpenActivitiesToExcel = () => {
    if (selectedPersonsForActivities.length === 0) return;
    const selectedSet = new Set(selectedPersonsForActivities.map(id => String(id)));
    const filteredActivities = getFilteredActivities().filter(a =>
      a.personId && selectedSet.has(String(a.personId))
    );
    if (filteredActivities.length === 0) {
      alert(t.noActivitiesFound);
      return;
    }

    const excelData = filteredActivities.map((activity, index) => ({
      [t.rowNo]: index + 1,
      [t.name]: activity.personName,
      [t.kids]: activity.personKidsNumber,
      [t.season]: activity.seasonName || t.uncategorizedSeason,
      [t.activity]: activity.description || '-',
      [t.distributor]: activity.distributor || '-'
    }));

    const worksheet = XLSX.utils.json_to_sheet(excelData);
    finalizeExcelWorksheet(worksheet, { narrowColumnIndexes: [0] });
    const workbook = XLSX.utils.book_new();
    applyWorkbookRtl(workbook);
    XLSX.utils.book_append_sheet(workbook, worksheet, t.openActivities);
    const fileName = `Selected_Open_Activities_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  };

  // Toggle activity status (close/open activity)
  const toggleActivityStatusInTable = async (personId, activityIndex) => {
    try {
      const person = persons.find(p => p._id === personId);
      if (!person || !person.activities || !person.activities[activityIndex]) return;

      const activity = person.activities[activityIndex];
      const currentStatus = activity.status || 'pending';
      const newStatus = currentStatus === 'pending' ? 'completed' : 'pending';

      // Show confirmation only when marking as completed
      if (newStatus === 'completed') {
        const activityDescription = activity.description || 'this activity';
        if (!window.confirm(t.confirmCompleteActivity.replace('{activity}', activityDescription))) {
          return; // User cancelled
        }
      }

      const updatedActivities = person.activities.map((act, idx) => {
        if (idx === activityIndex) {
          return {
            ...act,
            status: newStatus
          };
        }
        return act;
      });

      const response = await fetch(`${API_URL}/${personId}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(preparePersonBody({
          ...person,
          activities: updatedActivities
        })),
      });

      if (response.ok) {
        fetchPersons(); // Refresh the list
      } else {
        alert(t.errorUpdatingActivity);
      }
    } catch (error) {
      console.error('Error updating activity status:', error);
      alert(t.errorUpdatingActivity);
    }
  };

  // Update activity distributor
  const updateActivityDistributor = async (personId, activityIndex, distributorName) => {
    try {
      const person = persons.find(p => p._id === personId);
      if (!person) return;

      const updatedActivities = person.activities.map((activity, index) => {
        if (index === activityIndex) {
          return {
            ...activity,
            distributor: distributorName
          };
        }
        return activity;
      });

      const response = await fetch(`${API_URL}/${personId}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(preparePersonBody({
          ...person,
          activities: updatedActivities
        })),
      });

      if (response.ok) {
        fetchPersons(); // Refresh the list
      } else {
        alert(t.errorUpdatingDistributor);
      }
    } catch (error) {
      console.error('Error updating distributor:', error);
      alert(t.errorUpdatingDistributor);
    }
  };

  const updateActivitySeason = async (personId, activityIndex, seasonIdOrNull) => {
    try {
      const person = persons.find(p => p._id === personId);
      if (!person) return;

      const updatedActivities = person.activities.map((activity, index) => {
        if (index === activityIndex) {
          return {
            ...activity,
            season: seasonIdOrNull || null
          };
        }
        return activity;
      });

      const response = await fetch(`${API_URL}/${personId}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(preparePersonBody({
          ...person,
          activities: updatedActivities
        })),
      });

      if (response.ok) {
        fetchPersons();
      } else {
        alert(t.errorUpdatingSeason || t.errorUpdatingActivity);
      }
    } catch (error) {
      console.error('Error updating activity season:', error);
      alert(t.errorUpdatingSeason || t.errorUpdatingActivity);
    }
  };

  const handleAdd = () => {
    setEditingPerson(null);
    setFormData({
      name: '',
      phone: '',
      kids: [],
      monthIncome: '',
      maritalStatus: 'Single',
      liveInRenta: null,
      hasCar: null,
      bankNumber: '',
      description: ''
    });
    setModalVisible(true);
  };

  const handleEdit = (person) => {
    setEditingPerson(person);
    setFormData({
      name: person.name,
      phone: person.phone,
      kids: person.kids && person.kids.length > 0
        ? person.kids.map(kid => ({
            name: kid.name || '',
            birthday: kid.birthday ? formatDateForInput(kid.birthday) : ''
          }))
        : [],
      monthIncome: person.monthIncome != null && person.monthIncome !== '' ? person.monthIncome.toString() : '',
      maritalStatus: person.maritalStatus,
      liveInRenta: person.liveInRenta !== undefined && person.liveInRenta !== null ? person.liveInRenta : null,
      hasCar: person.hasCar !== undefined && person.hasCar !== null ? person.hasCar : null,
      bankNumber: person.bankNumber || '',
      description: person.description || ''
    });
    setModalVisible(true);
  };

  const handleSave = async () => {
    try {
      // Validate and format kids data
      const validKids = formData.kids
        .map(kid => ({
          name: kid.name.trim(),
          birthday: kid.birthday ? new Date(kid.birthday).toISOString() : null
        }))
        .filter(kid => kid.name && kid.birthday);

      const dataToSave = {
        name: formData.name,
        phone: formData.phone,
        kids: validKids,
        monthIncome: formData.monthIncome.trim() === '' ? null : (parseFloat(formData.monthIncome) || 0),
        maritalStatus: formData.maritalStatus,
        liveInRenta: formData.liveInRenta === null || formData.liveInRenta === undefined ? null : !!formData.liveInRenta,
        hasCar: formData.hasCar === null || formData.hasCar === undefined ? null : !!formData.hasCar,
        bankNumber: (formData.bankNumber || '').trim(),
        description: (formData.description || '').trim() || null
      };

      const url = editingPerson ? `${API_URL}/${editingPerson._id}` : API_URL;
      const method = editingPerson ? 'PUT' : 'POST';

      const response = await fetch(url, {
        method,
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(dataToSave),
      });

      if (response.ok) {
        setModalVisible(false);
        fetchPersons();
      } else {
        const error = await response.json();
        const message = (response.status === 409 && error.message === 'DUPLICATE_PERSON')
          ? (t.duplicatePerson || 'A person with this name and phone already exists.')
          : (error.message || error.description || 'Failed to save');
        alert('Error: ' + message);
      }
    } catch (error) {
      console.error('Error saving person:', error);
      alert('Error saving person');
    }
  };

  // Add a new kid
  const addKid = () => {
    setFormData({
      ...formData,
      kids: [...formData.kids, { name: '', birthday: '' }]
    });
  };

  // Remove a kid
  const removeKid = (index) => {
    const newKids = formData.kids.filter((_, i) => i !== index);
    setFormData({
      ...formData,
      kids: newKids
    });
  };

  // Update a kid's field
  const updateKid = (index, field, value) => {
    const newKids = [...formData.kids];
    newKids[index] = { ...newKids[index], [field]: value };
    setFormData({
      ...formData,
      kids: newKids
    });
  };

  // Update one part of kid's birthday (day, month, year) and rebuild full birthday.
  // Always persist all three parts (day, month, year) so clearing one part doesn't clear the others.
  const updateKidBirthdayPart = (index, part, value) => {
    const kid = formData.kids[index];
    const parts = getBirthdayParts(kid);
    const day = part === 'day' ? value : (kid.birthdayDay !== undefined ? kid.birthdayDay : parts.day);
    const month = part === 'month' ? value : (kid.birthdayMonth !== undefined ? kid.birthdayMonth : parts.month);
    const year = part === 'year' ? value : (kid.birthdayYear !== undefined ? kid.birthdayYear : parts.year);
    const newBirthday = buildBirthdayFromParts(day, month, year);
    const newKids = [...formData.kids];
    newKids[index] = {
      ...kid,
      birthdayDay: day,
      birthdayMonth: month,
      birthdayYear: year,
      birthday: newBirthday
    };
    setFormData({ ...formData, kids: newKids });
  };

  // Display value for one birthday part (raw input or parsed from birthday)
  const getKidBirthdayPart = (kid, part) => {
    const raw = part === 'day' ? kid.birthdayDay : part === 'month' ? kid.birthdayMonth : kid.birthdayYear;
    if (raw !== undefined && raw !== '') return raw;
    const parts = getBirthdayParts(kid);
    return parts[part];
  };

  // Toggle kids visibility for a person
  const toggleKidsVisibility = (personId) => {
    setExpandedKids(prev => ({
      ...prev,
      [personId]: !prev[personId]
    }));
  };

  // Open activities modal for a person
  const openActivitiesModal = (person) => {
    setSelectedPersonForActivities(person);
    setNewActivity('');
    setNewActivitySeasonId(seasons[0]?._id || '');
    setActivitySeasonDropdownOpen(false);
    setActivitiesModalVisible(true);
  };

  // Close activities modal
  const closeActivitiesModal = () => {
    setActivitiesModalVisible(false);
    setSelectedPersonForActivities(null);
    setNewActivity('');
    setActivitySeasonDropdownOpen(false);
  };

  // Add new activity
  const addActivity = async () => {
    if (!newActivity.trim() || !selectedPersonForActivities) return;
    if (seasons.length === 0) {
      alert(t.addSeasonFirst);
      return;
    }
    if (!newActivitySeasonId) {
      alert(t.selectSeason);
      return;
    }

    try {
      const updatedActivities = [
        ...(selectedPersonForActivities.activities || []),
        {
          description: newActivity.trim(),
          date: new Date().toISOString(),
          status: 'pending',
          distributor: '',
          season: newActivitySeasonId
        }
      ];

      const response = await fetch(`${API_URL}/${selectedPersonForActivities._id}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(preparePersonBody({
          ...selectedPersonForActivities,
          activities: updatedActivities
        })),
      });

      if (response.ok) {
        const updatedPerson = await response.json();
        setSelectedPersonForActivities(updatedPerson);
        setNewActivity('');
        fetchPersons(); // Refresh the list
      } else {
        const error = await response.json();
        alert(t.errorAddingActivity + ': ' + (error.message || ''));
      }
    } catch (error) {
      console.error('Error adding activity:', error);
      alert(t.errorAddingActivity);
    }
  };

  // Add activity to selected persons (with optional distributor)
  const addActivityToSelected = async (activityDescription, distributorName = '', seasonId = '') => {
    if (!activityDescription.trim() || selectedPersonsForActivities.length === 0) {
      alert(t.selectPeopleFirst || 'Please select at least one person first');
      return;
    }
    if (seasons.length === 0) {
      alert(t.addSeasonFirst);
      return;
    }
    if (!seasonId) {
      alert(t.selectSeason);
      return;
    }

    try {
      const promises = selectedPersonsForActivities.map(async (personId) => {
        const person = persons.find(p => p._id === personId);
        if (!person) return;

        const updatedActivities = [
          ...(person.activities || []),
          {
            description: activityDescription.trim(),
            date: new Date().toISOString(),
            status: 'pending',
            distributor: distributorName || '',
            season: seasonId
          }
        ];

        const response = await fetch(`${API_URL}/${personId}`, {
          method: 'PUT',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(preparePersonBody({
            ...person,
            activities: updatedActivities
          })),
        });

        if (!response.ok) {
          throw new Error(`Failed to update person ${person.name}`);
        }
      });

      await Promise.all(promises);
      setSelectedPersonsForActivities([]);
      setBulkAddActivityModalVisible(false);
      setBulkAddActivityDescription('');
      setBulkAddActivityDistributor('');
      setBulkAddActivitySeasonId(seasons[0]?._id || '');
      setBulkAddSeasonDropdownOpen(false);
      fetchPersons(); // Refresh the list
      alert(t.activityAddedToSelected || `Activity added to ${selectedPersonsForActivities.length} person(s)`);
    } catch (error) {
      console.error('Error adding activity to selected persons:', error);
      alert(t.errorAddingActivity);
    }
  };

  // Export selected persons to Excel (main page)
  const exportSelectedPersonsToExcel = () => {
    if (selectedPersonsForActivities.length === 0) return;
    const selectedPersons = persons.filter(p => selectedPersonsForActivities.includes(p._id));
    if (selectedPersons.length === 0) return;

    const excelData = selectedPersons.map(person => ({
      [t.name]: person.name || '',
      [t.phone]: person.phone || '',
      [t.kids]: person.kidsNumber ?? (person.kids?.length ?? 0),
      [t.income]: person.monthIncome != null && person.monthIncome !== '' ? person.monthIncome : '',
      [t.status]: t[person.maritalStatus?.toLowerCase()] || person.maritalStatus || '',
      [t.renta]: person.liveInRenta === true ? t.yes : person.liveInRenta === false ? t.no : (t.noAvailableData || 'No available data'),
      [t.car]: person.hasCar === true ? t.yes : person.hasCar === false ? t.no : (t.noAvailableData || 'No available data'),
      [t.bank]: person.bankNumber || '',
      [t.description]: person.description || ''
    }));

    const worksheet = XLSX.utils.json_to_sheet(excelData);
    finalizeExcelWorksheet(worksheet, { narrowColumnIndexes: [] });
    const workbook = XLSX.utils.book_new();
    applyWorkbookRtl(workbook);
    XLSX.utils.book_append_sheet(workbook, worksheet, t.peopleManagement || 'People');
    const fileName = `Selected_Persons_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  };

  // Toggle activity status
  const toggleActivityStatus = async (activityIndex) => {
    if (!selectedPersonForActivities) return;

    const activity = selectedPersonForActivities.activities[activityIndex];
    if (!activity) {
      console.error('Activity not found at index:', activityIndex);
      return;
    }

    const currentStatus = activity.status || 'pending';
    const newStatus = currentStatus === 'pending' ? 'completed' : 'pending';
    const activityDescription = activity.description || 'this activity';

    console.log('Toggling activity status:', {
      activityIndex,
      currentStatus,
      newStatus,
      activity
    });

    // Show confirmation only when marking as completed
    if (newStatus === 'completed') {
      if (!window.confirm(t.confirmCompleteActivity.replace('{activity}', activityDescription))) {
        return; // User cancelled
      }
    }

    try {
      const updatedActivities = selectedPersonForActivities.activities.map((act, idx) => {
        if (idx === activityIndex) {
          return {
            ...act,
            status: newStatus
          };
        }
        return act;
      });

      console.log('Updated activities:', updatedActivities);

      const response = await fetch(`${API_URL}/${selectedPersonForActivities._id}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(preparePersonBody({
          ...selectedPersonForActivities,
          activities: updatedActivities
        })),
      });

      if (response.ok) {
        const updatedPerson = await response.json();
        console.log('Updated person from server:', updatedPerson);
        setSelectedPersonForActivities(updatedPerson);
        fetchPersons(); // Refresh the list
      } else {
        const error = await response.json();
        console.error('Error response:', error);
        alert('Error updating activity status: ' + (error.message || ''));
      }
    } catch (error) {
      console.error('Error updating activity status:', error);
      alert('Error updating activity status');
    }
  };

  // Remove activity
  const removeActivity = async (activityIndex) => {
    if (!selectedPersonForActivities) return;

    // Get activity description for confirmation message
    const activity = selectedPersonForActivities.activities[activityIndex];
    const activityDescription = activity?.description || 'this activity';

    // Show confirmation dialog
    if (!window.confirm(`Are you sure you want to remove "${activityDescription}"?`)) {
      return; // User cancelled
    }

    try {
      const updatedActivities = selectedPersonForActivities.activities.filter(
        (_, index) => index !== activityIndex
      );

      const response = await fetch(`${API_URL}/${selectedPersonForActivities._id}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(preparePersonBody({
          ...selectedPersonForActivities,
          activities: updatedActivities
        })),
      });

      if (response.ok) {
        const updatedPerson = await response.json();
        setSelectedPersonForActivities(updatedPerson);
        fetchPersons(); // Refresh the list
      } else {
        const error = await response.json();
        alert(t.errorRemovingActivity + ': ' + (error.message || ''));
      }
    } catch (error) {
      console.error('Error removing activity:', error);
      alert(t.errorRemovingActivity);
    }
  };

  const handleDelete = async (id) => {
    if (window.confirm('Are you sure you want to delete this person?')) {
      try {
        const response = await fetch(`${API_URL}/${id}`, {
          method: 'DELETE',
        });

        if (response.ok) {
          fetchPersons();
        } else {
          alert(t.errorDeleting);
        }
      } catch (error) {
        console.error('Error deleting person:', error);
        alert(t.errorDeleting);
      }
    }
  };

  const toggleFavorite = async (personId) => {
    try {
      const person = persons.find(p => p._id === personId);
      if (!person) return;

      const updatedPerson = { ...person, favorite: !person.favorite };

      const response = await fetch(`${API_URL}/${personId}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(preparePersonBody(updatedPerson)),
      });

      if (response.ok) {
        const saved = await response.json();
        setPersons(prev =>
          prev.map(p => (p._id === personId ? saved : p))
        );
      } else {
        const error = await response.json();
        alert((t.errorUpdatingFavorite || 'Error updating favorite') + ': ' + (error.message || ''));
      }
    } catch (error) {
      console.error('Error updating favorite:', error);
      alert(t.errorUpdatingFavorite || 'Error updating favorite');
    }
  };

  const addSeason = async () => {
    const name = newSeasonName.trim();
    if (!name) {
      alert(t.seasonNameRequired);
      return;
    }
    try {
      const response = await fetch(SEASONS_API, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name })
      });
      if (response.ok) {
        setNewSeasonName('');
        await fetchSeasons();
      } else {
        const err = await response.json().catch(() => ({}));
        const message = err.description || err.message || t.errorSavingSeason;
        alert(message);
      }
    } catch (error) {
      console.error('Error adding season:', error);
      alert(t.errorSavingSeason);
    }
  };

  const removeSeason = async (seasonId) => {
    if (!window.confirm(t.confirmRemoveSeason)) return;
    try {
      const response = await fetch(`${SEASONS_API}/${seasonId}`, { method: 'DELETE' });
      if (response.ok) {
        await fetchSeasons();
      } else {
        const err = await response.json().catch(() => ({}));
        alert(err.description || err.message || t.errorRemovingSeason);
      }
    } catch (error) {
      console.error('Error removing season:', error);
      alert(t.errorRemovingSeason);
    }
  };

  const renderBulkAddActivityModal = () => (
    <Modal
      animationType="slide"
      transparent={true}
      visible={bulkAddActivityModalVisible}
      onRequestClose={() => {
        setBulkAddActivityModalVisible(false);
        setBulkAddActivityDescription('');
        setBulkAddActivityDistributor('');
        setBulkAddDistributorDropdownOpen(false);
        setBulkAddSeasonDropdownOpen(false);
      }}
    >
      <View style={styles.modalOverlay}>
        <View style={[styles.modalContent, isRTL && styles.modalContentRTL]}>
          <Text style={styles.modalTitle}>
            {t.addActivityToSelected || 'Add Activity'} ({selectedPersonsForActivities.length})
          </Text>
          <View style={[styles.form, styles.bulkAddForm]}>
            <View style={styles.formGroup}>
              <Text style={styles.label}>{t.activityDescription}</Text>
              <TextInput
                style={[styles.input, isRTL && styles.inputRTL]}
                value={bulkAddActivityDescription}
                onChangeText={setBulkAddActivityDescription}
                placeholder={t.activityDescription}
                multiline
                textAlign={isRTL ? 'right' : 'left'}
              />
            </View>
            <View style={[styles.formGroup, styles.bulkAddDistributorFormGroup]}>
              <Text style={styles.label}>{t.distributor}</Text>
              <TouchableOpacity
                style={styles.distributorSelectButton}
                onPress={() => setBulkAddDistributorDropdownOpen(!bulkAddDistributorDropdownOpen)}
              >
                <Text style={styles.distributorSelectButtonText}>
                  {bulkAddActivityDistributor || t.selectDistributor}
                </Text>
                <Text style={styles.distributorSelectArrow}>
                  {bulkAddDistributorDropdownOpen ? '▲' : '▼'}
                </Text>
              </TouchableOpacity>
              {bulkAddDistributorDropdownOpen && (
                <View style={styles.distributorDropdown}>
                  <ScrollView style={styles.distributorDropdownScroll} nestedScrollEnabled={true}>
                    <TouchableOpacity
                      style={styles.distributorOption}
                      onPress={() => {
                        setBulkAddActivityDistributor('');
                        setBulkAddDistributorDropdownOpen(false);
                      }}
                    >
                      <Text style={styles.distributorOptionText}>{t.noDistributor}</Text>
                    </TouchableOpacity>
                    {distributors.map(distributor => (
                      <TouchableOpacity
                        key={distributor.id}
                        style={styles.distributorOption}
                        onPress={() => {
                          setBulkAddActivityDistributor(distributor.name);
                          setBulkAddDistributorDropdownOpen(false);
                        }}
                      >
                        <Text style={styles.distributorOptionText}>
                          {distributor.name} ({distributor.phone})
                        </Text>
                      </TouchableOpacity>
                    ))}
                  </ScrollView>
                </View>
              )}
            </View>
            <View style={[styles.formGroup, styles.bulkAddDistributorFormGroup]}>
              <Text style={styles.label}>{t.selectSeason}</Text>
              <TouchableOpacity
                style={styles.distributorSelectButton}
                onPress={() => setBulkAddSeasonDropdownOpen(!bulkAddSeasonDropdownOpen)}
              >
                <Text style={styles.distributorSelectButtonText}>
                  {bulkAddActivitySeasonId && seasons.find(s => String(s._id) === String(bulkAddActivitySeasonId))
                    ? seasons.find(s => String(s._id) === String(bulkAddActivitySeasonId)).name
                    : t.selectSeason}
                </Text>
                <Text style={styles.distributorSelectArrow}>
                  {bulkAddSeasonDropdownOpen ? '▲' : '▼'}
                </Text>
              </TouchableOpacity>
              {bulkAddSeasonDropdownOpen && (
                <View style={styles.distributorDropdown}>
                  <ScrollView style={styles.distributorDropdownScroll} nestedScrollEnabled={true}>
                    {seasons.map(season => (
                      <TouchableOpacity
                        key={season._id}
                        style={styles.distributorOption}
                        onPress={() => {
                          setBulkAddActivitySeasonId(season._id);
                          setBulkAddSeasonDropdownOpen(false);
                        }}
                      >
                        <Text style={styles.distributorOptionText}>{season.name}</Text>
                      </TouchableOpacity>
                    ))}
                  </ScrollView>
                </View>
              )}
            </View>
            <View style={[styles.formGroup, { flexDirection: isRTL ? 'row-reverse' : 'row', gap: 10, marginTop: 16 }]}>
              <TouchableOpacity
                style={[styles.cancelButton, { flex: 1 }]}
                onPress={() => {
                  setBulkAddActivityModalVisible(false);
                  setBulkAddActivityDescription('');
                  setBulkAddActivityDistributor('');
                  setBulkAddDistributorDropdownOpen(false);
                  setBulkAddSeasonDropdownOpen(false);
                }}
              >
                <Text style={styles.cancelButtonText}>{t.close}</Text>
              </TouchableOpacity>
              <TouchableOpacity
                style={[styles.addButton, { flex: 1 }]}
                onPress={() => {
                  if (bulkAddActivityDescription.trim()) {
                    addActivityToSelected(
                      bulkAddActivityDescription.trim(),
                      bulkAddActivityDistributor,
                      bulkAddActivitySeasonId
                    );
                  } else {
                    alert(t.activityDescription || 'Enter activity description');
                  }
                }}
              >
                <Text style={styles.addButtonText}>{t.add}</Text>
              </TouchableOpacity>
            </View>
          </View>
        </View>
      </View>
    </Modal>
  );

  const checkboxHeaderCellStyle = {
    width: 50,
    flex: 0,
    minWidth: 50,
    maxWidth: 50,
    justifyContent: 'center',
    alignItems: 'center',
    paddingHorizontal: 8
  };

  // Render Open Activities page
  const renderOpenActivitiesPage = () => {
    const filteredActivities = getFilteredActivities();
    const activityGroups = groupActivitiesBySeason(filteredActivities);
    const visiblePersonIds = [...new Set(filteredActivities.map(a => a.personId).filter(Boolean))];
    const allVisiblePeopleSelected =
      visiblePersonIds.length > 0 &&
      visiblePersonIds.every(pid =>
        selectedPersonsForActivities.some(sid => String(sid) === String(pid))
      );

    return (
      <View style={[styles.container, isRTL && styles.containerRTL]}>
        <View style={[styles.header, isRTL && styles.headerRTL]}>
          <Text style={styles.title}>{t.openActivities}</Text>
          <View style={styles.headerRight}>
            {selectedPersonsForActivities.length > 0 && (
              <>
                <TouchableOpacity
                  style={styles.exportSelectedExcelButton}
                  onPress={exportSelectedOpenActivitiesToExcel}
                >
                  <Text style={styles.exportSelectedExcelButtonText}>
                    {t.exportSelectedToExcel || 'Export to Excel'} ({selectedPersonsForActivities.length})
                  </Text>
                </TouchableOpacity>
                <TouchableOpacity
                  style={styles.bulkActivityButton}
                  onPress={() => {
                    setBulkAddActivityModalVisible(true);
                    if (seasons.length && !bulkAddActivitySeasonId) {
                      setBulkAddActivitySeasonId(seasons[0]._id);
                    }
                  }}
                >
                  <Text style={styles.bulkActivityButtonText}>
                    {t.addActivityToSelected || 'Add Activity'} ({selectedPersonsForActivities.length})
                  </Text>
                </TouchableOpacity>
              </>
            )}
            <TouchableOpacity 
              style={styles.exportExcelButton}
              onPress={exportToExcel}
            >
              <Text style={styles.exportExcelButtonText}>{t.exportToExcel}</Text>
            </TouchableOpacity>
            <TouchableOpacity 
              style={styles.manageDistributorsButton}
              onPress={() => setDistributorModalVisible(true)}
            >
              <Text style={styles.manageDistributorsButtonText}>{t.manageDistributors}</Text>
            </TouchableOpacity>
            <TouchableOpacity 
              style={styles.manageDistributorsButton}
              onPress={() => setSeasonModalVisible(true)}
            >
              <Text style={styles.manageDistributorsButtonText}>{t.manageSeasons}</Text>
            </TouchableOpacity>
            <TouchableOpacity 
              style={styles.navButton}
              onPress={() => {
                setSelectedPersonsForActivities([]);
                setCurrentView('main');
              }}
            >
              <Text style={styles.navButtonText}>{t.backToMain}</Text>
            </TouchableOpacity>
            <TouchableOpacity 
              style={styles.languageButton}
              onPress={() => setLanguage(language === 'en' ? 'ar' : 'en')}
            >
              <Text style={styles.languageButtonText}>{language === 'en' ? 'عربي' : 'English'}</Text>
            </TouchableOpacity>
          </View>
        </View>

        {/* Filters Section */}
        <View style={styles.filtersContainer}>
          <TouchableOpacity 
            style={[styles.filtersHeader, isRTL && styles.filtersHeaderRTL]}
            onPress={() => setShowOpenActivitiesFilters(!showOpenActivitiesFilters)}
          >
            <Text style={[styles.filtersHeaderText, isRTL && styles.filtersHeaderTextRTL]}>
              {t.filters} {showOpenActivitiesFilters ? '▼' : (isRTL ? '◀' : '▶')}
            </Text>
            <TouchableOpacity
              style={styles.clearFiltersButton}
              onPress={() => {
                setOpenActivitiesFilters({
                  distributor: '',
                  season: '',
                  activityStatus: 'open',
                  activityName: '',
                  dateFromDay: '',
                  dateFromMonth: '',
                  dateFromYear: '',
                  dateToDay: '',
                  dateToMonth: '',
                  dateToYear: ''
                });
              }}
            >
              <Text style={styles.clearFiltersButtonText}>{t.clearAll}</Text>
            </TouchableOpacity>
          </TouchableOpacity>

          {showOpenActivitiesFilters && (
            <View style={[styles.filtersContent, isRTL && styles.filtersContentRTL]}>
              <View style={[styles.filterRow, isRTL && styles.filterRowRTL, { zIndex: 9999999 }]}>
                <View style={[styles.filterGroup, { zIndex: 9999999 }]}>
                  <Text style={[styles.filterLabel, isRTL && styles.filterLabelRTL]}>{t.distributor}</Text>
                  <View style={styles.distributorSelector}>
                    <TouchableOpacity
                      style={styles.distributorSelectButton}
                      onPress={() => {
                        const dropdownKey = 'filter-distributor';
                        setOpenDropdowns(prev => ({
                          ...prev,
                          [dropdownKey]: !prev[dropdownKey]
                        }));
                      }}
                    >
                      <Text style={styles.distributorSelectButtonText}>
                        {openActivitiesFilters.distributor === '__NO_DISTRIBUTOR__' 
                          ? t.noDistributor 
                          : openActivitiesFilters.distributor || t.all}
                      </Text>
                      <Text style={styles.distributorSelectArrow}>
                        {openDropdowns['filter-distributor'] ? '▲' : '▼'}
                      </Text>
                    </TouchableOpacity>
                    {openDropdowns['filter-distributor'] && (
                      <View style={styles.distributorDropdown}>
                        <ScrollView 
                          style={styles.distributorDropdownScroll}
                          nestedScrollEnabled={true}
                        >
                          <TouchableOpacity
                            style={styles.distributorOption}
                            onPress={() => {
                              setOpenActivitiesFilters(prev => ({ ...prev, distributor: '' }));
                              setOpenDropdowns(prev => ({ ...prev, 'filter-distributor': false }));
                            }}
                          >
                            <Text style={styles.distributorOptionText}>{t.all}</Text>
                          </TouchableOpacity>
                          <TouchableOpacity
                            style={styles.distributorOption}
                            onPress={() => {
                              setOpenActivitiesFilters(prev => ({ ...prev, distributor: '__NO_DISTRIBUTOR__' }));
                              setOpenDropdowns(prev => ({ ...prev, 'filter-distributor': false }));
                            }}
                          >
                            <Text style={styles.distributorOptionText}>{t.noDistributor}</Text>
                          </TouchableOpacity>
                          {distributors.map(distributor => (
                            <TouchableOpacity
                              key={distributor.id}
                              style={styles.distributorOption}
                              onPress={() => {
                                setOpenActivitiesFilters(prev => ({ ...prev, distributor: distributor.name }));
                                setOpenDropdowns(prev => ({ ...prev, 'filter-distributor': false }));
                              }}
                            >
                              <Text style={styles.distributorOptionText}>
                                {distributor.name} ({distributor.phone})
                              </Text>
                            </TouchableOpacity>
                          ))}
                        </ScrollView>
                      </View>
                    )}
                  </View>
                </View>
                <View style={styles.filterGroup}>
                  <Text style={[styles.filterLabel, isRTL && styles.filterLabelRTL]}>{t.stillActivityOpen}</Text>
                  <View style={[styles.filterRadioGroup, isRTL && styles.filterRadioGroupRTL]}>
                    <TouchableOpacity
                      style={[
                        styles.filterRadioOption,
                        openActivitiesFilters.activityStatus === 'all' && styles.filterRadioOptionSelected
                      ]}
                      onPress={() => setOpenActivitiesFilters(prev => ({ ...prev, activityStatus: 'all' }))}
                    >
                      <Text style={[
                        styles.filterRadioText,
                        openActivitiesFilters.activityStatus === 'all' && styles.filterRadioTextSelected
                      ]}>{t.all}</Text>
                    </TouchableOpacity>
                    <TouchableOpacity
                      style={[
                        styles.filterRadioOption,
                        openActivitiesFilters.activityStatus === 'open' && styles.filterRadioOptionSelected
                      ]}
                      onPress={() => setOpenActivitiesFilters(prev => ({ ...prev, activityStatus: 'open' }))}
                    >
                      <Text style={[
                        styles.filterRadioText,
                        openActivitiesFilters.activityStatus === 'open' && styles.filterRadioTextSelected
                      ]}>{t.open}</Text>
                    </TouchableOpacity>
                    <TouchableOpacity
                      style={[
                        styles.filterRadioOption,
                        openActivitiesFilters.activityStatus === 'closed' && styles.filterRadioOptionSelected
                      ]}
                      onPress={() => setOpenActivitiesFilters(prev => ({ ...prev, activityStatus: 'closed' }))}
                    >
                      <Text style={[
                        styles.filterRadioText,
                        openActivitiesFilters.activityStatus === 'closed' && styles.filterRadioTextSelected
                      ]}>{t.closed}</Text>
                    </TouchableOpacity>
                  </View>
                </View>
              </View>
              <View style={[styles.filterRow, isRTL && styles.filterRowRTL, { zIndex: 8888888 }]}>
                <View style={[styles.filterGroup, { zIndex: 8888888 }]}>
                  <Text style={[styles.filterLabel, isRTL && styles.filterLabelRTL]}>{t.season}</Text>
                  <View style={styles.distributorSelector}>
                    <TouchableOpacity
                      style={styles.distributorSelectButton}
                      onPress={() => {
                        const dropdownKey = 'filter-season';
                        setOpenDropdowns(prev => ({
                          ...prev,
                          [dropdownKey]: !prev[dropdownKey]
                        }));
                      }}
                    >
                      <Text style={styles.distributorSelectButtonText}>
                        {openActivitiesFilters.season === '__NO_SEASON__'
                          ? t.noSeasonAssigned
                          : openActivitiesFilters.season && seasons.find(s => String(s._id) === String(openActivitiesFilters.season))
                            ? seasons.find(s => String(s._id) === String(openActivitiesFilters.season)).name
                            : t.all}
                      </Text>
                      <Text style={styles.distributorSelectArrow}>
                        {openDropdowns['filter-season'] ? '▲' : '▼'}
                      </Text>
                    </TouchableOpacity>
                    {openDropdowns['filter-season'] && (
                      <View style={styles.distributorDropdown}>
                        <ScrollView
                          style={styles.distributorDropdownScroll}
                          nestedScrollEnabled={true}
                        >
                          <TouchableOpacity
                            style={styles.distributorOption}
                            onPress={() => {
                              setOpenActivitiesFilters(prev => ({ ...prev, season: '' }));
                              setOpenDropdowns(prev => ({ ...prev, 'filter-season': false }));
                            }}
                          >
                            <Text style={styles.distributorOptionText}>{t.all}</Text>
                          </TouchableOpacity>
                          <TouchableOpacity
                            style={styles.distributorOption}
                            onPress={() => {
                              setOpenActivitiesFilters(prev => ({ ...prev, season: '__NO_SEASON__' }));
                              setOpenDropdowns(prev => ({ ...prev, 'filter-season': false }));
                            }}
                          >
                            <Text style={styles.distributorOptionText}>{t.noSeasonAssigned}</Text>
                          </TouchableOpacity>
                          {seasons.map(season => (
                            <TouchableOpacity
                              key={season._id}
                              style={styles.distributorOption}
                              onPress={() => {
                                setOpenActivitiesFilters(prev => ({ ...prev, season: String(season._id) }));
                                setOpenDropdowns(prev => ({ ...prev, 'filter-season': false }));
                              }}
                            >
                              <Text style={styles.distributorOptionText}>{season.name}</Text>
                            </TouchableOpacity>
                          ))}
                        </ScrollView>
                      </View>
                    )}
                  </View>
                </View>
                <View style={[styles.filterGroup, { zIndex: 1 }]}>
                  <Text style={[styles.filterLabel, isRTL && styles.filterLabelRTL]}>{t.activityName}</Text>
                  <TextInput
                    style={[styles.filterInput, isRTL && styles.filterInputRTL]}
                    value={openActivitiesFilters.activityName}
                    onChangeText={(text) => setOpenActivitiesFilters(prev => ({ ...prev, activityName: text }))}
                    placeholder={t.searchActivityName}
                    textAlign={isRTL ? 'right' : 'left'}
                  />
                </View>
              </View>
              <View style={[styles.filterRow, isRTL && styles.filterRowRTL]}>
                <View style={[styles.filterGroup, styles.timelineFilterGroup]}>
                  <Text style={[styles.filterLabel, isRTL && styles.filterLabelRTL]}>{t.timelineFrom}</Text>
                  <View style={styles.birthdayRow}>
                    <TextInput
                      style={[styles.birthdayInput, isRTL && styles.inputRTL]}
                      value={openActivitiesFilters.dateFromDay}
                      onChangeText={(text) =>
                        setOpenActivitiesFilters(prev => ({
                          ...prev,
                          dateFromDay: text.replace(/\D/g, '').slice(0, 2)
                        }))
                      }
                      placeholder={t.day}
                      keyboardType="number-pad"
                      textAlign={isRTL ? 'right' : 'left'}
                    />
                    <TextInput
                      style={[styles.birthdayInput, isRTL && styles.inputRTL]}
                      value={openActivitiesFilters.dateFromMonth}
                      onChangeText={(text) =>
                        setOpenActivitiesFilters(prev => ({
                          ...prev,
                          dateFromMonth: text.replace(/\D/g, '').slice(0, 2)
                        }))
                      }
                      placeholder={t.month}
                      keyboardType="number-pad"
                      textAlign={isRTL ? 'right' : 'left'}
                    />
                    <TextInput
                      style={[styles.birthdayInput, isRTL && styles.inputRTL]}
                      value={openActivitiesFilters.dateFromYear}
                      onChangeText={(text) =>
                        setOpenActivitiesFilters(prev => ({
                          ...prev,
                          dateFromYear: text.replace(/\D/g, '').slice(0, 4)
                        }))
                      }
                      placeholder={t.year}
                      keyboardType="number-pad"
                      textAlign={isRTL ? 'right' : 'left'}
                    />
                  </View>
                </View>
                <View style={[styles.filterGroup, styles.timelineFilterGroup]}>
                  <Text style={[styles.filterLabel, isRTL && styles.filterLabelRTL]}>{t.timelineTo}</Text>
                  <View style={styles.birthdayRow}>
                    <TextInput
                      style={[styles.birthdayInput, isRTL && styles.inputRTL]}
                      value={openActivitiesFilters.dateToDay}
                      onChangeText={(text) =>
                        setOpenActivitiesFilters(prev => ({
                          ...prev,
                          dateToDay: text.replace(/\D/g, '').slice(0, 2)
                        }))
                      }
                      placeholder={t.day}
                      keyboardType="number-pad"
                      textAlign={isRTL ? 'right' : 'left'}
                    />
                    <TextInput
                      style={[styles.birthdayInput, isRTL && styles.inputRTL]}
                      value={openActivitiesFilters.dateToMonth}
                      onChangeText={(text) =>
                        setOpenActivitiesFilters(prev => ({
                          ...prev,
                          dateToMonth: text.replace(/\D/g, '').slice(0, 2)
                        }))
                      }
                      placeholder={t.month}
                      keyboardType="number-pad"
                      textAlign={isRTL ? 'right' : 'left'}
                    />
                    <TextInput
                      style={[styles.birthdayInput, isRTL && styles.inputRTL]}
                      value={openActivitiesFilters.dateToYear}
                      onChangeText={(text) =>
                        setOpenActivitiesFilters(prev => ({
                          ...prev,
                          dateToYear: text.replace(/\D/g, '').slice(0, 4)
                        }))
                      }
                      placeholder={t.year}
                      keyboardType="number-pad"
                      textAlign={isRTL ? 'right' : 'left'}
                    />
                  </View>
                </View>
              </View>
            </View>
          )}
        </View>

        <ScrollView 
          style={styles.scrollView}
          onScrollBeginDrag={() => setOpenDropdowns({})}
        >
          <View style={[styles.table, isRTL && styles.tableRTL]}>
            <View style={[styles.tableHeader, isRTL && styles.tableHeaderRTL]}>
              <View style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL, checkboxHeaderCellStyle]}>
                <TouchableOpacity
                  style={[
                    styles.checkbox,
                    allVisiblePeopleSelected && visiblePersonIds.length > 0 && styles.checkboxChecked
                  ]}
                  onPress={() => {
                    if (allVisiblePeopleSelected) {
                      setSelectedPersonsForActivities(prev =>
                        prev.filter(id => !visiblePersonIds.some(vid => String(vid) === String(id)))
                      );
                    } else {
                      setSelectedPersonsForActivities(prev => {
                        const next = [...prev];
                        visiblePersonIds.forEach(vid => {
                          if (!next.some(id => String(id) === String(vid))) {
                            next.push(vid);
                          }
                        });
                        return next;
                      });
                    }
                  }}
                  disabled={visiblePersonIds.length === 0}
                >
                  {allVisiblePeopleSelected && visiblePersonIds.length > 0 && (
                    <Text style={styles.checkboxCheckmark}>✓</Text>
                  )}
                </TouchableOpacity>
              </View>
              <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL, { flex: 2 }]}>{t.name}</Text>
              <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL]}>{t.kids}</Text>
              <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL, { flex: 2 }]}>{t.activity}</Text>
              <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL, { flex: 1.4 }]}>{t.season}</Text>
              <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL, { flex: 2 }]}>{t.distributor}</Text>
              <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL]}>{t.status}</Text>
            </View>

            {filteredActivities.length === 0 ? (
              <View style={styles.emptyState}>
                <Text style={styles.emptyText}>{t.noActivitiesFound}</Text>
              </View>
            ) : (
              activityGroups.map((group) => (
                <View key={group.key}>
                  <View style={[styles.seasonSectionHeader, isRTL && styles.seasonSectionHeaderRTL]}>
                    <Text style={[styles.seasonSectionHeaderText, isRTL && styles.seasonSectionHeaderTextRTL]}>{group.title}</Text>
                  </View>
                  {group.activities.map((activity, index) => {
                const isPending = (activity.activityStatus === 'pending' || !activity.activityStatus) && activity.description;
                const hasNoActivity = !activity.description || activity.activityStatus === 'none';
                
                const distDropdownKey = `${activity.personId}-${activity.activityIndex}`;
                const seasonDropdownKey = `season-${activity.personId}-${activity.activityIndex}`;
                const isDistOpen = openDropdowns[distDropdownKey];
                const isSeasonOpen = openDropdowns[seasonDropdownKey];
                const isDropdownOpen = isDistOpen || isSeasonOpen;
                const isPersonSelected = selectedPersonsForActivities.some(
                  sid => String(sid) === String(activity.personId)
                );
                return (
                  <View key={`${activity.personId}-${activity.date}-${group.key}-${index}`} style={[styles.tableRow, isRTL && styles.tableRowRTL, { zIndex: isDropdownOpen ? 99999999 : 1 }]}>
                    <View style={[styles.tableCell, checkboxHeaderCellStyle]}>
                      <TouchableOpacity
                        style={[styles.checkbox, isPersonSelected && styles.checkboxChecked]}
                        onPress={() => {
                          if (isPersonSelected) {
                            setSelectedPersonsForActivities(prev =>
                              prev.filter(id => String(id) !== String(activity.personId))
                            );
                          } else {
                            setSelectedPersonsForActivities(prev => [...prev, activity.personId]);
                          }
                        }}
                      >
                        {isPersonSelected && <Text style={styles.checkboxCheckmark}>✓</Text>}
                      </TouchableOpacity>
                    </View>
                    <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL, { flex: 2 }]}>
                      {activity.personName}
                    </Text>
                    <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>
                      {activity.personKidsNumber}
                    </Text>
                    <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL, { flex: 2 }]}>
                      {activity.description || '-'}
                    </Text>
                  <View style={[styles.tableCell, { flex: 1.4, zIndex: isSeasonOpen ? 99999999 : 1 }]}>
                    {activity.activityStatus !== 'none' ? (
                      <View style={[styles.distributorSelector, { zIndex: isSeasonOpen ? 99999999 : 1 }]}>
                        <TouchableOpacity
                          style={styles.distributorSelectButton}
                          onPress={(e) => {
                            e.stopPropagation();
                            setOpenDropdowns(prev => ({
                              ...prev,
                              [seasonDropdownKey]: !prev[seasonDropdownKey]
                            }));
                          }}
                        >
                          <Text style={styles.distributorSelectButtonText} numberOfLines={1}>
                            {activity.seasonName || t.noSeasonAssigned}
                          </Text>
                          <Text style={styles.distributorSelectArrow}>
                            {openDropdowns[seasonDropdownKey] ? '▲' : '▼'}
                          </Text>
                        </TouchableOpacity>
                        {openDropdowns[seasonDropdownKey] && (
                          <View style={[styles.distributorDropdown, { zIndex: 99999999 }]}>
                            <ScrollView 
                              style={styles.distributorDropdownScroll}
                              nestedScrollEnabled={true}
                            >
                              <TouchableOpacity
                                style={styles.distributorOption}
                                onPress={(e) => {
                                  e.stopPropagation();
                                  updateActivitySeason(activity.personId, activity.activityIndex, null);
                                  setOpenDropdowns(prev => ({
                                    ...prev,
                                    [seasonDropdownKey]: false
                                  }));
                                }}
                              >
                                <Text style={styles.distributorOptionText}>{t.noSeasonAssigned}</Text>
                              </TouchableOpacity>
                              {seasons.map(season => (
                                <TouchableOpacity
                                  key={season._id}
                                  style={styles.distributorOption}
                                  onPress={(e) => {
                                    e.stopPropagation();
                                    updateActivitySeason(activity.personId, activity.activityIndex, season._id);
                                    setOpenDropdowns(prev => ({
                                      ...prev,
                                      [seasonDropdownKey]: false
                                    }));
                                  }}
                                >
                                  <Text style={styles.distributorOptionText}>{season.name}</Text>
                                </TouchableOpacity>
                              ))}
                            </ScrollView>
                          </View>
                        )}
                      </View>
                    ) : (
                      <Text style={[styles.tableCellText, isRTL && styles.tableCellRTL]}>-</Text>
                    )}
                  </View>
                  <View style={[styles.tableCell, { flex: 2, zIndex: isDistOpen ? 99999999 : 1 }]}>
                    {activity.activityStatus !== 'none' ? (
                      <View style={[styles.distributorSelector, { zIndex: isDistOpen ? 99999999 : 1 }]}>
                        <TouchableOpacity
                          style={styles.distributorSelectButton}
                          onPress={(e) => {
                            e.stopPropagation();
                            setOpenDropdowns(prev => ({
                              ...prev,
                              [distDropdownKey]: !prev[distDropdownKey]
                            }));
                          }}
                        >
                          <Text style={styles.distributorSelectButtonText}>
                            {activity.distributor || t.selectDistributor}
                          </Text>
                          <Text style={styles.distributorSelectArrow}>
                            {openDropdowns[distDropdownKey] ? '▲' : '▼'}
                          </Text>
                        </TouchableOpacity>
                        {openDropdowns[distDropdownKey] && (
                          <View style={[styles.distributorDropdown, { zIndex: 99999999 }]}>
                            <ScrollView 
                              style={styles.distributorDropdownScroll}
                              nestedScrollEnabled={true}
                            >
                              <TouchableOpacity
                                style={styles.distributorOption}
                                onPress={(e) => {
                                  e.stopPropagation();
                                  updateActivityDistributor(activity.personId, activity.activityIndex, '');
                                  setOpenDropdowns(prev => ({
                                    ...prev,
                                    [distDropdownKey]: false
                                  }));
                                }}
                              >
                                <Text style={styles.distributorOptionText}>{t.selectDistributor}</Text>
                              </TouchableOpacity>
                              {distributors.map(distributor => (
                                <TouchableOpacity
                                  key={distributor.id}
                                  style={styles.distributorOption}
                                  onPress={(e) => {
                                    e.stopPropagation();
                                    updateActivityDistributor(activity.personId, activity.activityIndex, distributor.name);
                                    setOpenDropdowns(prev => ({
                                      ...prev,
                                      [distDropdownKey]: false
                                    }));
                                  }}
                                >
                                  <Text style={styles.distributorOptionText}>
                                    {distributor.name} ({distributor.phone})
                                  </Text>
                                </TouchableOpacity>
                              ))}
                            </ScrollView>
                          </View>
                        )}
                      </View>
                    ) : (
                      <Text style={[styles.tableCellText, isRTL && styles.tableCellRTL]}>-</Text>
                    )}
                  </View>
                  <View style={[styles.tableCell]}>
                    {hasNoActivity ? (
                      <Text style={[styles.tableCellText, isRTL && styles.tableCellRTL]}>{t.na}</Text>
                    ) : isPending ? (
                      <TouchableOpacity
                        style={styles.statusButtonPending}
                        onPress={() => {
                          if (activity.activityIndex !== undefined) {
                            toggleActivityStatusInTable(activity.personId, activity.activityIndex);
                          }
                        }}
                      >
                        <Text style={styles.statusButtonText}>{t.pending}</Text>
                      </TouchableOpacity>
                    ) : (
                      <Text style={[styles.tableCellText, isRTL && styles.tableCellRTL, styles.statusCompleted]}>
                        {t.completed}
                      </Text>
                    )}
                  </View>
                </View>
                );
                  })}
                </View>
              ))
            )}
          </View>
        </ScrollView>

        {/* Distributor Management Modal */}
        <Modal
          animationType="slide"
          transparent={true}
          visible={distributorModalVisible}
          onRequestClose={() => setDistributorModalVisible(false)}
        >
          <View style={styles.modalOverlay}>
            <View style={[styles.modalContent, isRTL && styles.modalContentRTL]}>
              <Text style={styles.modalTitle}>{t.manageDistributors}</Text>

              <ScrollView style={styles.form}>
                {/* Add New Distributor */}
                <View style={styles.formGroup}>
                  <Text style={styles.label}>{t.addNewDistributor}</Text>
                  <TextInput
                    style={[styles.input, isRTL && styles.inputRTL]}
                    value={newDistributorName}
                    onChangeText={setNewDistributorName}
                    placeholder={t.distributorName}
                    textAlign={isRTL ? 'right' : 'left'}
                  />
                  <TextInput
                    style={[styles.input, isRTL && styles.inputRTL, { marginTop: 8 }]}
                    value={newDistributorPhone}
                    onChangeText={setNewDistributorPhone}
                    placeholder={t.distributorPhone}
                    keyboardType="phone-pad"
                    textAlign={isRTL ? 'right' : 'left'}
                  />
                  <TouchableOpacity
                    style={styles.addDistributorButton}
                    onPress={addDistributor}
                  >
                    <Text style={styles.addDistributorButtonText}>{t.add}</Text>
                  </TouchableOpacity>
                </View>

                {/* Distributors List */}
                <View style={styles.formGroup}>
                  <Text style={styles.label}>{t.distributorsList} ({distributors.length})</Text>
                  {distributors.length === 0 ? (
                    <Text style={styles.hintText}>{t.noDistributors}</Text>
                  ) : (
                    distributors.map(distributor => (
                      <View key={distributor.id} style={styles.distributorItem}>
                        <View style={styles.distributorItemInfo}>
                          <Text style={styles.distributorItemName}>{distributor.name}</Text>
                          <Text style={styles.distributorItemPhone}>{distributor.phone}</Text>
                        </View>
                        <TouchableOpacity
                          style={styles.removeDistributorButton}
                          onPress={() => removeDistributor(distributor.id)}
                        >
                          <Text style={styles.removeDistributorButtonText}>{t.remove}</Text>
                        </TouchableOpacity>
                      </View>
                    ))
                  )}
                </View>
              </ScrollView>

              <View style={[styles.modalButtons, isRTL && styles.modalButtonsRTL]}>
                <TouchableOpacity
                  style={[styles.modalButton, styles.cancelButton]}
                  onPress={() => {
                    setDistributorModalVisible(false);
                    setNewDistributorName('');
                    setNewDistributorPhone('');
                  }}
                >
                  <Text style={styles.cancelButtonText}>{t.close}</Text>
                </TouchableOpacity>
              </View>
            </View>
          </View>
        </Modal>

        {/* Season Management Modal */}
        <Modal
          animationType="slide"
          transparent={true}
          visible={seasonModalVisible}
          onRequestClose={() => setSeasonModalVisible(false)}
        >
          <View style={styles.modalOverlay}>
            <View style={[styles.modalContent, isRTL && styles.modalContentRTL]}>
              <Text style={styles.modalTitle}>{t.manageSeasons}</Text>

              <ScrollView style={styles.form}>
                <View style={styles.formGroup}>
                  <Text style={styles.label}>{t.addNewSeason}</Text>
                  <TextInput
                    style={[styles.input, isRTL && styles.inputRTL]}
                    value={newSeasonName}
                    onChangeText={setNewSeasonName}
                    placeholder={t.seasonNamePlaceholder}
                    textAlign={isRTL ? 'right' : 'left'}
                  />
                  <TouchableOpacity
                    style={styles.addDistributorButton}
                    onPress={addSeason}
                  >
                    <Text style={styles.addDistributorButtonText}>{t.add}</Text>
                  </TouchableOpacity>
                </View>

                <View style={styles.formGroup}>
                  <Text style={styles.label}>{t.seasonsList} ({seasons.length})</Text>
                  {seasons.length === 0 ? (
                    <Text style={styles.hintText}>{t.noSeasons}</Text>
                  ) : (
                    seasons.map(season => (
                      <View key={season._id} style={styles.distributorItem}>
                        <View style={styles.distributorItemInfo}>
                          <Text style={styles.distributorItemName}>{season.name}</Text>
                        </View>
                        <TouchableOpacity
                          style={styles.removeDistributorButton}
                          onPress={() => removeSeason(season._id)}
                        >
                          <Text style={styles.removeDistributorButtonText}>{t.remove}</Text>
                        </TouchableOpacity>
                      </View>
                    ))
                  )}
                </View>
              </ScrollView>

              <View style={[styles.modalButtons, isRTL && styles.modalButtonsRTL]}>
                <TouchableOpacity
                  style={[styles.modalButton, styles.cancelButton]}
                  onPress={() => {
                    setSeasonModalVisible(false);
                    setNewSeasonName('');
                  }}
                >
                  <Text style={styles.cancelButtonText}>{t.close}</Text>
                </TouchableOpacity>
              </View>
            </View>
          </View>
        </Modal>

        {renderBulkAddActivityModal()}
      </View>
    );
  };

  // Render main page or Open Activities page
  if (currentView === 'openActivities') {
    return renderOpenActivitiesPage();
  }

  return (
    <View style={[styles.container, isRTL && styles.containerRTL]}>
      <View style={[styles.header, isRTL && styles.headerRTL]}>
        <Text style={styles.title}>{t.title}</Text>
        <View style={styles.headerRight}>
          {selectedPersonsForActivities.length > 0 && (
            <>
              <TouchableOpacity
                style={styles.exportSelectedExcelButton}
                onPress={exportSelectedPersonsToExcel}
              >
                <Text style={styles.exportSelectedExcelButtonText}>
                  {t.exportSelectedToExcel || 'Export to Excel'} ({selectedPersonsForActivities.length})
                </Text>
              </TouchableOpacity>
              <TouchableOpacity 
                style={styles.bulkActivityButton}
                onPress={() => {
                  setBulkAddActivityModalVisible(true);
                  if (seasons.length && !bulkAddActivitySeasonId) {
                    setBulkAddActivitySeasonId(seasons[0]._id);
                  }
                }}
              >
                <Text style={styles.bulkActivityButtonText}>
                  {t.addActivityToSelected || 'Add Activity'} ({selectedPersonsForActivities.length})
                </Text>
              </TouchableOpacity>
            </>
          )}
          <TouchableOpacity 
            style={styles.navButton}
            onPress={() => {
              fetchPersons();
              fetchSeasons();
              setSelectedPersonsForActivities([]);
              setCurrentView('openActivities');
            }}
          >
            <Text style={styles.navButtonText}>{t.openActivities}</Text>
          </TouchableOpacity>
          <TouchableOpacity 
            style={styles.languageButton}
            onPress={() => setLanguage(language === 'en' ? 'ar' : 'en')}
          >
            <Text style={styles.languageButtonText}>{language === 'en' ? 'عربي' : 'English'}</Text>
          </TouchableOpacity>
          <TouchableOpacity style={styles.addButton} onPress={handleAdd}>
            <Text style={styles.addButtonText}>{t.addPerson}</Text>
          </TouchableOpacity>
        </View>
      </View>

      {renderBulkAddActivityModal()}

      {/* Search Bar */}
      <View style={[styles.searchContainer, isRTL && styles.searchContainerRTL]}>
        <TextInput
          style={[styles.searchInput, isRTL && styles.searchInputRTL]}
          placeholder={t.searchPlaceholder}
          value={searchQuery}
          onChangeText={setSearchQuery}
          placeholderTextColor="#999"
          textAlign={isRTL ? 'right' : 'left'}
        />
        {searchQuery.length > 0 && (
          <TouchableOpacity 
            style={styles.clearSearchButton}
            onPress={() => setSearchQuery('')}
          >
            <Text style={styles.clearSearchButtonText}>✕</Text>
          </TouchableOpacity>
        )}
      </View>

      {/* Favorites Filter */}
      <View style={[styles.favoritesFilterContainer, isRTL && styles.favoritesFilterContainerRTL]}>
        <TouchableOpacity
          style={[
            styles.favoritesFilterButton,
            !showFavoritesOnly && styles.favoritesFilterButtonActive
          ]}
          onPress={() => setShowFavoritesOnly(false)}
        >
          <Text
            style={[
              styles.favoritesFilterButtonText,
              !showFavoritesOnly && styles.favoritesFilterButtonTextActive
            ]}
          >
            {t.favoritesFilterAll}
          </Text>
        </TouchableOpacity>
        <TouchableOpacity
          style={[
            styles.favoritesFilterButton,
            showFavoritesOnly && styles.favoritesFilterButtonActive
          ]}
          onPress={() => setShowFavoritesOnly(true)}
        >
          <Text
            style={[
              styles.favoritesFilterButtonText,
              showFavoritesOnly && styles.favoritesFilterButtonTextActive
            ]}
          >
            {t.favoritesFilterFavorites}
          </Text>
        </TouchableOpacity>
      </View>

      {/* Filters Section */}
      <View style={styles.filtersContainer}>
        <TouchableOpacity 
          style={[styles.filtersHeader, isRTL && styles.filtersHeaderRTL]}
          onPress={() => setShowFilters(!showFilters)}
        >
          <Text style={[styles.filtersHeaderText, isRTL && styles.filtersHeaderTextRTL]}>
            {t.filters} {showFilters ? '▼' : (isRTL ? '◀' : '▶')}
          </Text>
          <TouchableOpacity
            style={styles.clearFiltersButton}
            onPress={() => {
              setFilters({
                minIncome: '',
                maxIncome: '',
                liveInRenta: null,
                hasCar: null,
                maritalStatus: null,
                kidsMinAge: '',
                kidsMaxAge: '',
                minKidsNumber: ''
              });
            }}
          >
            <Text style={styles.clearFiltersButtonText}>{t.clearAll}</Text>
          </TouchableOpacity>
        </TouchableOpacity>

        {showFilters && (
          <View style={[styles.filtersContent, isRTL && styles.filtersContentRTL]}>
            <View style={[styles.filterRow, isRTL && styles.filterRowRTL]}>
              <View style={styles.filterGroup}>
                <Text style={[styles.filterLabel, isRTL && styles.filterLabelRTL]}>{t.minIncome}</Text>
                <TextInput
                  style={[styles.filterInput, isRTL && styles.filterInputRTL]}
                  value={filters.minIncome}
                  onChangeText={(text) => setFilters({ ...filters, minIncome: text })}
                  placeholder="0"
                  keyboardType="numeric"
                  textAlign={isRTL ? 'right' : 'left'}
                />
              </View>
              <View style={styles.filterGroup}>
                <Text style={[styles.filterLabel, isRTL && styles.filterLabelRTL]}>{t.maxIncome}</Text>
                <TextInput
                  style={[styles.filterInput, isRTL && styles.filterInputRTL]}
                  value={filters.maxIncome}
                  onChangeText={(text) => setFilters({ ...filters, maxIncome: text })}
                  placeholder="999999"
                  keyboardType="numeric"
                  textAlign={isRTL ? 'right' : 'left'}
                />
              </View>
            </View>

            <View style={[styles.filterRow, isRTL && styles.filterRowRTL]}>
              <View style={styles.filterGroup}>
                <Text style={[styles.filterLabel, isRTL && styles.filterLabelRTL]}>{t.liveInRenta}</Text>
                <View style={[styles.filterRadioGroup, isRTL && styles.filterRadioGroupRTL]}>
                  <TouchableOpacity
                    style={[
                      styles.filterRadioOption,
                      filters.liveInRenta === null && styles.filterRadioOptionSelected
                    ]}
                    onPress={() => setFilters({ ...filters, liveInRenta: null })}
                  >
                    <Text style={[
                      styles.filterRadioText,
                      filters.liveInRenta === null && styles.filterRadioTextSelected
                    ]}>{t.all}</Text>
                  </TouchableOpacity>
                  <TouchableOpacity
                    style={[
                      styles.filterRadioOption,
                      filters.liveInRenta === true && styles.filterRadioOptionSelected
                    ]}
                    onPress={() => setFilters({ ...filters, liveInRenta: true })}
                  >
                    <Text style={[
                      styles.filterRadioText,
                      filters.liveInRenta === true && styles.filterRadioTextSelected
                    ]}>{t.yes}</Text>
                  </TouchableOpacity>
                  <TouchableOpacity
                    style={[
                      styles.filterRadioOption,
                      filters.liveInRenta === false && styles.filterRadioOptionSelected
                    ]}
                    onPress={() => setFilters({ ...filters, liveInRenta: false })}
                  >
                    <Text style={[
                      styles.filterRadioText,
                      filters.liveInRenta === false && styles.filterRadioTextSelected
                    ]}>{t.no}</Text>
                  </TouchableOpacity>
                </View>
              </View>
              <View style={styles.filterGroup}>
                <Text style={[styles.filterLabel, isRTL && styles.filterLabelRTL]}>{t.hasCar}</Text>
                <View style={[styles.filterRadioGroup, isRTL && styles.filterRadioGroupRTL]}>
                  <TouchableOpacity
                    style={[
                      styles.filterRadioOption,
                      filters.hasCar === null && styles.filterRadioOptionSelected
                    ]}
                    onPress={() => setFilters({ ...filters, hasCar: null })}
                  >
                    <Text style={[
                      styles.filterRadioText,
                      filters.hasCar === null && styles.filterRadioTextSelected
                    ]}>{t.all}</Text>
                  </TouchableOpacity>
                  <TouchableOpacity
                    style={[
                      styles.filterRadioOption,
                      filters.hasCar === true && styles.filterRadioOptionSelected
                    ]}
                    onPress={() => setFilters({ ...filters, hasCar: true })}
                  >
                    <Text style={[
                      styles.filterRadioText,
                      filters.hasCar === true && styles.filterRadioTextSelected
                    ]}>{t.yes}</Text>
                  </TouchableOpacity>
                  <TouchableOpacity
                    style={[
                      styles.filterRadioOption,
                      filters.hasCar === false && styles.filterRadioOptionSelected
                    ]}
                    onPress={() => setFilters({ ...filters, hasCar: false })}
                  >
                    <Text style={[
                      styles.filterRadioText,
                      filters.hasCar === false && styles.filterRadioTextSelected
                    ]}>{t.no}</Text>
                  </TouchableOpacity>
                </View>
              </View>
            </View>

            <View style={[styles.filterRow, isRTL && styles.filterRowRTL]}>
              <View style={styles.filterGroup}>
                <Text style={[styles.filterLabel, isRTL && styles.filterLabelRTL]}>{t.maritalStatus}</Text>
                <View style={[styles.filterRadioGroup, isRTL && styles.filterRadioGroupRTL]}>
                  <TouchableOpacity
                    style={[
                      styles.filterRadioOption,
                      filters.maritalStatus === null && styles.filterRadioOptionSelected
                    ]}
                    onPress={() => setFilters({ ...filters, maritalStatus: null })}
                  >
                    <Text style={[
                      styles.filterRadioText,
                      filters.maritalStatus === null && styles.filterRadioTextSelected
                    ]}>{t.all}</Text>
                  </TouchableOpacity>
                  {['Single', 'Married', 'Divorced', 'Widowed', 'Divorce Project', 'Separated'].map((status) => (
                    <TouchableOpacity
                      key={status}
                      style={[
                        styles.filterRadioOption,
                        filters.maritalStatus === status && styles.filterRadioOptionSelected
                      ]}
                      onPress={() => setFilters({ ...filters, maritalStatus: status })}
                    >
                      <Text style={[
                        styles.filterRadioText,
                        filters.maritalStatus === status && styles.filterRadioTextSelected
                      ]}>
                        {t[status.toLowerCase()]}
                      </Text>
                    </TouchableOpacity>
                  ))}
                </View>
              </View>
            </View>

            <View style={[styles.filterRow, isRTL && styles.filterRowRTL]}>
              <View style={styles.filterGroup}>
                <Text style={[styles.filterLabel, isRTL && styles.filterLabelRTL]}>{t.kidsAgeRange}</Text>
                <View style={[styles.ageRangeContainer, isRTL && styles.ageRangeContainerRTL]}>
                  <TextInput
                    style={[styles.filterInput, { flex: 1, marginRight: isRTL ? 0 : 8, marginLeft: isRTL ? 8 : 0 }]}
                    value={filters.kidsMinAge}
                    onChangeText={(text) => setFilters({ ...filters, kidsMinAge: text })}
                    placeholder={t.minAge}
                    keyboardType="numeric"
                    textAlign={isRTL ? 'right' : 'left'}
                  />
                  <Text style={styles.ageRangeSeparator}>{isRTL ? 'إلى' : 'to'}</Text>
                  <TextInput
                    style={[styles.filterInput, { flex: 1, marginLeft: isRTL ? 0 : 8, marginRight: isRTL ? 8 : 0 }]}
                    value={filters.kidsMaxAge}
                    onChangeText={(text) => setFilters({ ...filters, kidsMaxAge: text })}
                    placeholder={t.maxAge}
                    keyboardType="numeric"
                    textAlign={isRTL ? 'right' : 'left'}
                  />
                </View>
              </View>
            </View>

            <View style={[styles.filterRow, isRTL && styles.filterRowRTL]}>
              <View style={styles.filterGroup}>
                <Text style={[styles.filterLabel, isRTL && styles.filterLabelRTL]}>{t.minNumberOfKids}</Text>
                <TextInput
                  style={[styles.filterInput, isRTL && styles.filterInputRTL]}
                  value={filters.minKidsNumber}
                  onChangeText={(text) => setFilters({ ...filters, minKidsNumber: text })}
                  placeholder="0"
                  keyboardType="numeric"
                  textAlign={isRTL ? 'right' : 'left'}
                />
              </View>
            </View>
          </View>
        )}
      </View>

      <ScrollView style={styles.scrollView}>
        <View style={[styles.table, isRTL && styles.tableRTL]}>
          <View style={[styles.tableHeader, isRTL && styles.tableHeaderRTL]}>
            <View style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL, { width: 50, flex: 0, minWidth: 50, maxWidth: 50, justifyContent: 'center', alignItems: 'center', paddingHorizontal: 8 }]}>
              <TouchableOpacity
                style={[
                  styles.checkbox,
                  filteredPersons.length > 0 && filteredPersons.every(p => selectedPersonsForActivities.includes(p._id)) && styles.checkboxChecked
                ]}
                onPress={() => {
                  const filteredIds = filteredPersons.map(p => p._id);
                  const allFilteredSelected = filteredPersons.length > 0 && filteredPersons.every(p => selectedPersonsForActivities.includes(p._id));
                  if (allFilteredSelected) {
                    setSelectedPersonsForActivities(prev => prev.filter(id => !filteredIds.includes(id)));
                  } else {
                    const combined = [...new Set([...selectedPersonsForActivities, ...filteredIds])];
                    setSelectedPersonsForActivities(combined);
                  }
                }}
                disabled={filteredPersons.length === 0}
              >
                {filteredPersons.length > 0 && filteredPersons.every(p => selectedPersonsForActivities.includes(p._id)) && (
                  <Text style={styles.checkboxCheckmark}>✓</Text>
                )}
              </TouchableOpacity>
            </View>
            <View style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL, { width: 50, flex: 0, minWidth: 50, maxWidth: 50, justifyContent: 'center', alignItems: 'center', paddingHorizontal: 8 }]}>
              <Text style={styles.favoriteHeaderIcon}>★</Text>
            </View>
            <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL, { flex: 2 }]}>{t.name}</Text>
            <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL]}>{t.phone}</Text>
            <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL]}>{t.kids}</Text>
            <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL]}>{t.income}</Text>
            <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL]}>{t.status}</Text>
            <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL]}>{t.renta}</Text>
            <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL]}>{t.car}</Text>
            <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL]}>{t.bank}</Text>
            <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL, { flex: 1.5 }]}>{t.actions}</Text>
          </View>

          {paginatedPersons.length === 0 && filteredPersons.length > 0 && (
            <View style={styles.emptyState}>
              <Text style={styles.emptyText}>{t.noResultsOnPage}</Text>
            </View>
          )}

          {filteredPersons.length === 0 && searchQuery.trim() && (
            <View style={styles.emptyState}>
              <Text style={styles.emptyText}>{t.noPeopleFound.replace('{query}', searchQuery)}</Text>
            </View>
          )}

          {filteredPersons.length === 0 && !searchQuery.trim() && (
            <View style={styles.emptyState}>
              <Text style={styles.emptyText}>{t.noPeopleAdded}</Text>
            </View>
          )}

          {paginatedPersons.map((person) => {
            const hasKids = person.kids && person.kids.length > 0;
            const hasDescription = person.description && String(person.description).trim();
            const hasExpandableDetails = hasKids || hasDescription;
            const isExpanded = expandedKids[person._id];
            const isSelected = selectedPersonsForActivities.includes(person._id);
            const isFavorite = !!person.favorite;
            
            return (
              <View key={person._id}>
                <View style={[styles.tableRow, isRTL && styles.tableRowRTL]}>
                  <View style={[styles.tableCell, { width: 50, flex: 0, minWidth: 50, maxWidth: 50, justifyContent: 'center', alignItems: 'center', paddingHorizontal: 8 }]}>
                    <TouchableOpacity
                      style={[styles.checkbox, isSelected && styles.checkboxChecked]}
                      onPress={() => {
                        if (isSelected) {
                          setSelectedPersonsForActivities(selectedPersonsForActivities.filter(id => id !== person._id));
                        } else {
                          setSelectedPersonsForActivities([...selectedPersonsForActivities, person._id]);
                        }
                      }}
                    >
                      {isSelected && <Text style={styles.checkboxCheckmark}>✓</Text>}
                    </TouchableOpacity>
                  </View>
                  <View style={[styles.tableCell, { width: 50, flex: 0, minWidth: 50, maxWidth: 50, justifyContent: 'center', alignItems: 'center', paddingHorizontal: 8 }]}>
                    <TouchableOpacity onPress={() => toggleFavorite(person._id)}>
                      <Text style={[styles.favoriteIcon, isFavorite && styles.favoriteIconActive]}>
                        {isFavorite ? '★' : '☆'}
                      </Text>
                    </TouchableOpacity>
                  </View>
                  <View style={[styles.tableCell, { flex: 2, flexDirection: isRTL ? 'row-reverse' : 'row', alignItems: 'center' }]}>
                    {hasExpandableDetails && (
                      <TouchableOpacity 
                        style={styles.expandButton}
                        onPress={() => toggleKidsVisibility(person._id)}
                      >
                        <Text style={styles.expandIcon}>
                          {isExpanded ? '▼' : (isRTL ? '◀' : '▶')}
                        </Text>
                      </TouchableOpacity>
                    )}
                    <Text style={[styles.tableCellText, { flex: 1, textAlign: isRTL ? 'right' : 'left' }]}>{person.name}</Text>
                  </View>
                  <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>{person.phone}</Text>
                  <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>{person.kidsNumber || 0}</Text>
                  <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>{person.monthIncome != null && person.monthIncome !== '' ? person.monthIncome : ''}</Text>
                  <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>{t[person.maritalStatus.toLowerCase()] || person.maritalStatus}</Text>
                  <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>{person.liveInRenta === true ? t.yes : person.liveInRenta === false ? t.no : (t.noAvailableData || 'No available data')}</Text>
                  <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>{person.hasCar === true ? t.yes : person.hasCar === false ? t.no : (t.noAvailableData || 'No available data')}</Text>
                  <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>{person.bankNumber || ''}</Text>
                  <View style={[styles.tableCell, styles.actionsCell]}>
                    <TouchableOpacity 
                      style={styles.activitiesButton} 
                      onPress={() => openActivitiesModal(person)}
                    >
                      <Text style={styles.activitiesButtonText}>📋</Text>
                    </TouchableOpacity>
                    <TouchableOpacity style={styles.editButton} onPress={() => handleEdit(person)}>
                      <Text style={styles.editButtonText}>{t.edit || 'Edit'}</Text>
                    </TouchableOpacity>
                    <TouchableOpacity style={styles.deleteButton} onPress={() => handleDelete(person._id)}>
                      <Text style={styles.deleteButtonText}>{t.delete || 'Delete'}</Text>
                    </TouchableOpacity>
                  </View>
                </View>
                {hasExpandableDetails && isExpanded && (
                  <View style={[styles.kidsDetailsRow, isRTL && styles.kidsDetailsRowRTL]}>
                    {hasKids && (
                      <View style={styles.detailSection}>
                        <Text style={[styles.kidsLabel, isRTL && styles.kidsLabelRTL]}>{t.kidsDetails}</Text>
                        {person.kids.map((kid, index) => {
                          const age = calculateAge(kid.birthday);
                          return (
                            <View key={index} style={styles.kidItem}>
                              <Text style={[styles.kidText, isRTL && styles.kidTextRTL]}>
                                {kid.name} - {t.age} {age !== null ? `${age} ${t.years}` : 'N/A'}
                                {kid.birthday && ` (${t.born} ${new Date(kid.birthday).toLocaleDateString()})`}
                              </Text>
                            </View>
                          );
                        })}
                      </View>
                    )}
                    {hasDescription && (
                      <View style={[styles.detailSection, hasKids && { marginTop: 8 }]}>
                        <Text style={[styles.kidsLabel, isRTL && styles.kidsLabelRTL]}>{t.description}</Text>
                        <Text style={[styles.kidText, isRTL && styles.kidTextRTL]}>{person.description.trim()}</Text>
                      </View>
                    )}
                  </View>
                )}
              </View>
            );
          })}

        </View>
      </ScrollView>

      {/* View Options and Pagination - Bottom of Page */}
      <View style={styles.paginationContainer}>
        <View style={styles.viewOptionsContainer}>
          <Text style={styles.viewOptionsLabel}>{t.view}</Text>
          {[15, 25, 50, 'All'].map((option) => (
            <TouchableOpacity
              key={option}
              style={[
                styles.viewOptionButton,
                itemsPerPage === option && styles.viewOptionButtonSelected
              ]}
              onPress={() => {
                setItemsPerPage(option);
                setCurrentPage(1);
              }}
            >
              <Text style={[
                styles.viewOptionText,
                itemsPerPage === option && styles.viewOptionTextSelected
              ]}>
                {option === 'All' ? t.all : option}
              </Text>
            </TouchableOpacity>
          ))}
        </View>
        {itemsPerPage !== 'All' && filteredPersons.length > parseInt(itemsPerPage) && (
          <View style={styles.paginationControls}>
            <Text style={styles.paginationInfo}>
              {t.showing} {((currentPage - 1) * parseInt(itemsPerPage)) + 1} - {Math.min(currentPage * parseInt(itemsPerPage), filteredPersons.length)} {t.of} {filteredPersons.length}
            </Text>
            <View style={styles.paginationButtons}>
              <TouchableOpacity
                style={[
                  styles.paginationButton,
                  currentPage === 1 && styles.paginationButtonDisabled
                ]}
                onPress={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                disabled={currentPage === 1}
              >
                <Text style={[
                  styles.paginationButtonText,
                  currentPage === 1 && styles.paginationButtonTextDisabled
                ]}>
                  {t.previous}
                </Text>
              </TouchableOpacity>
              <Text style={styles.pageNumber}>
                {t.page} {currentPage} {t.of} {Math.ceil(filteredPersons.length / parseInt(itemsPerPage))}
              </Text>
              <TouchableOpacity
                style={[
                  styles.paginationButton,
                  currentPage >= Math.ceil(filteredPersons.length / parseInt(itemsPerPage)) && styles.paginationButtonDisabled
                ]}
                onPress={() => setCurrentPage(prev => Math.min(Math.ceil(filteredPersons.length / parseInt(itemsPerPage)), prev + 1))}
                disabled={currentPage >= Math.ceil(filteredPersons.length / parseInt(itemsPerPage))}
              >
                <Text style={[
                  styles.paginationButtonText,
                  currentPage >= Math.ceil(filteredPersons.length / parseInt(itemsPerPage)) && styles.paginationButtonTextDisabled
                ]}>
                  {t.next}
                </Text>
              </TouchableOpacity>
            </View>
          </View>
        )}
        {itemsPerPage === 'All' && (
          <Text style={styles.paginationInfo}>
            {t.showingAll} {filteredPersons.length} {t.records}
          </Text>
        )}
      </View>

      <Modal
        animationType="slide"
        transparent={true}
        visible={modalVisible}
        onRequestClose={() => setModalVisible(false)}
      >
        <View style={styles.modalOverlay}>
          <View style={styles.modalContent}>
            <Text style={styles.modalTitle}>
              {editingPerson ? t.editPerson : t.addNewPerson}
            </Text>

            <ScrollView style={styles.form}>
              <View style={styles.formGroup}>
                <Text style={styles.label}>{t.nameField}</Text>
                <TextInput
                  style={[styles.input, isRTL && styles.inputRTL]}
                  value={formData.name}
                  onChangeText={(text) => setFormData({ ...formData, name: text })}
                  placeholder={isRTL ? 'أدخل الاسم' : 'Enter name'}
                  textAlign={isRTL ? 'right' : 'left'}
                />
              </View>

              <View style={styles.formGroup}>
                <Text style={styles.label}>{t.phoneField}</Text>
                <TextInput
                  style={[styles.input, isRTL && styles.inputRTL]}
                  value={formData.phone}
                  onChangeText={(text) => setFormData({ ...formData, phone: text })}
                  placeholder={isRTL ? 'أدخل رقم الهاتف' : 'Enter phone number'}
                  keyboardType="phone-pad"
                  textAlign={isRTL ? 'right' : 'left'}
                />
              </View>

              <View style={styles.formGroup}>
                <View style={[styles.kidsHeader, isRTL && styles.kidsHeaderRTL]}>
                  <Text style={styles.label}>{t.kidsDetailsField}</Text>
                  <TouchableOpacity style={styles.addKidButton} onPress={addKid}>
                    <Text style={styles.addKidButtonText}>{t.addKid}</Text>
                  </TouchableOpacity>
                </View>
                
                {formData.kids.map((kid, index) => (
                  <View key={index} style={styles.kidFormItem}>
                    <View style={[styles.kidFormHeader, isRTL && styles.kidFormHeaderRTL]}>
                      <Text style={styles.kidFormLabel}>{t.kidNumber} {index + 1}</Text>
                      <TouchableOpacity 
                        style={styles.removeKidButton} 
                        onPress={() => removeKid(index)}
                      >
                        <Text style={styles.removeKidButtonText}>{t.remove}</Text>
                      </TouchableOpacity>
                    </View>
                    <TextInput
                      style={[styles.input, isRTL && styles.inputRTL]}
                      value={kid.name}
                      onChangeText={(text) => updateKid(index, 'name', text)}
                      placeholder={t.kidsName}
                      textAlign={isRTL ? 'right' : 'left'}
                    />
                    <View style={styles.birthdayRow}>
                      <TextInput
                        style={[styles.birthdayInput, isRTL && styles.inputRTL]}
                        value={getKidBirthdayPart(kid, 'day')}
                        onChangeText={(text) => updateKidBirthdayPart(index, 'day', text.replace(/\D/g, '').slice(0, 2))}
                        placeholder={t.day}
                        keyboardType="number-pad"
                        textAlign={isRTL ? 'right' : 'left'}
                      />
                      <TextInput
                        style={[styles.birthdayInput, isRTL && styles.inputRTL]}
                        value={getKidBirthdayPart(kid, 'month')}
                        onChangeText={(text) => updateKidBirthdayPart(index, 'month', text.replace(/\D/g, '').slice(0, 2))}
                        placeholder={t.month}
                        keyboardType="number-pad"
                        textAlign={isRTL ? 'right' : 'left'}
                      />
                      <TextInput
                        style={[styles.birthdayInput, isRTL && styles.inputRTL]}
                        value={getKidBirthdayPart(kid, 'year')}
                        onChangeText={(text) => updateKidBirthdayPart(index, 'year', text.replace(/\D/g, '').slice(0, 4))}
                        placeholder={t.year}
                        keyboardType="number-pad"
                        textAlign={isRTL ? 'right' : 'left'}
                      />
                    </View>
                    {kid.birthday && calculateAge(kid.birthday) !== null && (
                      <Text style={styles.ageText}>
                        {t.age} {calculateAge(kid.birthday)} {t.years}
                      </Text>
                    )}
                  </View>
                ))}
                
                {formData.kids.length === 0 && (
                  <Text style={styles.hintText}>{t.clickAddKid}</Text>
                )}
                
                {formData.kids.length > 0 && (
                  <Text style={styles.kidsCountText}>
                    {t.totalKids} {formData.kids.length}
                  </Text>
                )}
              </View>

              <View style={styles.formGroup}>
                <Text style={styles.label}>{t.monthIncome}</Text>
                <TextInput
                  style={[styles.input, isRTL && styles.inputRTL]}
                  value={formData.monthIncome}
                  onChangeText={(text) => setFormData({ ...formData, monthIncome: text })}
                  placeholder={isRTL ? 'أدخل الدخل الشهري' : 'Enter monthly income'}
                  keyboardType="numeric"
                  textAlign={isRTL ? 'right' : 'left'}
                />
              </View>

              <View style={styles.formGroup}>
                <Text style={styles.label}>{t.maritalStatusField}</Text>
                <View style={styles.radioGroup}>
                  {['Single', 'Married', 'Divorced', 'Widowed', 'Divorce Project', 'Separated'].map((status) => (
                    <TouchableOpacity
                      key={status}
                      style={[
                        styles.radioOption,
                        formData.maritalStatus === status && styles.radioOptionSelected
                      ]}
                      onPress={() => setFormData({ ...formData, maritalStatus: status })}
                    >
                      <Text style={[
                        styles.radioText,
                        formData.maritalStatus === status && styles.radioTextSelected
                      ]}>
                        {t[status.toLowerCase()]}
                      </Text>
                    </TouchableOpacity>
                  ))}
                </View>
              </View>

              <View style={styles.formGroup}>
                <Text style={styles.label}>{t.liveInRenta}</Text>
                <View style={[styles.radioGroup, isRTL && styles.radioGroupRTL]}>
                  <TouchableOpacity
                    style={[styles.radioOption, formData.liveInRenta === true && styles.radioOptionSelected]}
                    onPress={() => setFormData({ ...formData, liveInRenta: true })}
                  >
                    <Text style={[styles.radioText, formData.liveInRenta === true && styles.radioTextSelected]}>{t.yes}</Text>
                  </TouchableOpacity>
                  <TouchableOpacity
                    style={[styles.radioOption, formData.liveInRenta === false && styles.radioOptionSelected]}
                    onPress={() => setFormData({ ...formData, liveInRenta: false })}
                  >
                    <Text style={[styles.radioText, formData.liveInRenta === false && styles.radioTextSelected]}>{t.no}</Text>
                  </TouchableOpacity>
                  <TouchableOpacity
                    style={[styles.radioOption, (formData.liveInRenta === null || formData.liveInRenta === undefined) && styles.radioOptionSelected]}
                    onPress={() => setFormData({ ...formData, liveInRenta: null })}
                  >
                    <Text style={[styles.radioText, (formData.liveInRenta === null || formData.liveInRenta === undefined) && styles.radioTextSelected]}>{t.noAvailableData}</Text>
                  </TouchableOpacity>
                </View>
              </View>

              <View style={styles.formGroup}>
                <Text style={styles.label}>{t.hasCar}</Text>
                <View style={[styles.radioGroup, isRTL && styles.radioGroupRTL]}>
                  <TouchableOpacity
                    style={[styles.radioOption, formData.hasCar === true && styles.radioOptionSelected]}
                    onPress={() => setFormData({ ...formData, hasCar: true })}
                  >
                    <Text style={[styles.radioText, formData.hasCar === true && styles.radioTextSelected]}>{t.yes}</Text>
                  </TouchableOpacity>
                  <TouchableOpacity
                    style={[styles.radioOption, formData.hasCar === false && styles.radioOptionSelected]}
                    onPress={() => setFormData({ ...formData, hasCar: false })}
                  >
                    <Text style={[styles.radioText, formData.hasCar === false && styles.radioTextSelected]}>{t.no}</Text>
                  </TouchableOpacity>
                  <TouchableOpacity
                    style={[styles.radioOption, (formData.hasCar === null || formData.hasCar === undefined) && styles.radioOptionSelected]}
                    onPress={() => setFormData({ ...formData, hasCar: null })}
                  >
                    <Text style={[styles.radioText, (formData.hasCar === null || formData.hasCar === undefined) && styles.radioTextSelected]}>{t.noAvailableData}</Text>
                  </TouchableOpacity>
                </View>
              </View>

              <View style={styles.formGroup}>
                <Text style={styles.label}>{t.bankNumber}</Text>
                <TextInput
                  style={[styles.input, isRTL && styles.inputRTL]}
                  value={formData.bankNumber}
                  onChangeText={(text) => setFormData({ ...formData, bankNumber: text })}
                  placeholder={isRTL ? 'أدخل رقم البنك' : 'Enter bank number'}
                  textAlign={isRTL ? 'right' : 'left'}
                />
              </View>

              <View style={styles.formGroup}>
                <Text style={styles.label}>{t.descriptionField}</Text>
                <TextInput
                  style={[styles.input, styles.textArea, isRTL && styles.inputRTL]}
                  value={formData.description}
                  onChangeText={(text) => setFormData({ ...formData, description: text })}
                  placeholder={t.descriptionPlaceholder}
                  textAlign={isRTL ? 'right' : 'left'}
                  multiline
                  numberOfLines={3}
                />
              </View>
            </ScrollView>

            <View style={[styles.modalButtons, isRTL && styles.modalButtonsRTL]}>
              <TouchableOpacity
                style={[styles.modalButton, styles.cancelButton]}
                onPress={() => setModalVisible(false)}
              >
                <Text style={styles.cancelButtonText}>{t.cancel}</Text>
              </TouchableOpacity>
              <TouchableOpacity
                style={[styles.modalButton, styles.saveButton]}
                onPress={handleSave}
              >
                <Text style={styles.saveButtonText}>{t.save}</Text>
              </TouchableOpacity>
            </View>
          </View>
        </View>
      </Modal>

      {/* Activities Modal */}
      <Modal
        animationType="slide"
        transparent={true}
        visible={activitiesModalVisible}
        onRequestClose={closeActivitiesModal}
      >
        <View style={styles.modalOverlay}>
          <View style={[styles.modalContent, isRTL && styles.modalContentRTL]}>
            <Text style={styles.modalTitle}>
              {t.activities} - {selectedPersonForActivities?.name || ''}
            </Text>

            <ScrollView style={styles.form}>
              {/* Add New Activity */}
              <View style={styles.formGroup}>
                <Text style={styles.label}>{t.selectSeason}</Text>
                <TouchableOpacity
                  style={styles.distributorSelectButton}
                  onPress={() => setActivitySeasonDropdownOpen(!activitySeasonDropdownOpen)}
                >
                  <Text style={styles.distributorSelectButtonText}>
                    {newActivitySeasonId && seasons.find(s => String(s._id) === String(newActivitySeasonId))
                      ? seasons.find(s => String(s._id) === String(newActivitySeasonId)).name
                      : t.selectSeason}
                  </Text>
                  <Text style={styles.distributorSelectArrow}>
                    {activitySeasonDropdownOpen ? '▲' : '▼'}
                  </Text>
                </TouchableOpacity>
                {activitySeasonDropdownOpen && (
                  <View style={[styles.distributorDropdown, { marginBottom: 8 }]}>
                    <ScrollView style={styles.distributorDropdownScroll} nestedScrollEnabled={true}>
                      {seasons.map(season => (
                        <TouchableOpacity
                          key={season._id}
                          style={styles.distributorOption}
                          onPress={() => {
                            setNewActivitySeasonId(season._id);
                            setActivitySeasonDropdownOpen(false);
                          }}
                        >
                          <Text style={styles.distributorOptionText}>{season.name}</Text>
                        </TouchableOpacity>
                      ))}
                    </ScrollView>
                  </View>
                )}
              </View>
              <View style={styles.formGroup}>
                <Text style={styles.label}>{t.addNewActivity}</Text>
                <View style={[styles.activityInputContainer, isRTL && styles.activityInputContainerRTL]}>
                  <TextInput
                    style={[styles.input, { flex: 1, marginRight: isRTL ? 0 : 8, marginLeft: isRTL ? 8 : 0 }]}
                    value={newActivity}
                    onChangeText={setNewActivity}
                    placeholder={t.activityDescription}
                    multiline
                    textAlign={isRTL ? 'right' : 'left'}
                  />
                  <TouchableOpacity 
                    style={styles.addActivityButton} 
                    onPress={addActivity}
                  >
                    <Text style={styles.addActivityButtonText}>{t.add}</Text>
                  </TouchableOpacity>
                </View>
              </View>

              {/* Activities List */}
              <View style={styles.formGroup}>
                <Text style={styles.label}>
                  {t.activityCount} ({selectedPersonForActivities?.activities?.length || 0})
                </Text>
                {selectedPersonForActivities?.activities && selectedPersonForActivities.activities.length > 0 ? (
                  [...selectedPersonForActivities.activities]
                    .map((activity, index) => ({ 
                      ...activity, 
                      _originalIndex: index,
                      status: activity.status || 'pending' // Ensure status exists
                    }))
                    .sort((a, b) => new Date(a.date) - new Date(b.date)) // Sort from oldest to newest
                    .map((activity) => {
                      const originalIndex = activity._originalIndex;
                      const currentStatus = activity.status || 'pending';
                      return (
                        <View key={`${activity.description}-${new Date(activity.date).getTime()}-${originalIndex}`} style={styles.activityItem}>
                          <View style={styles.activityContent}>
                            <Text style={[styles.activityDescription, isRTL && styles.activityDescriptionRTL]}>
                              {activity.description}
                            </Text>
                            <Text style={[styles.activityDate, isRTL && styles.activityDateRTL]}>
                              {new Date(activity.date).toLocaleString()}
                            </Text>
                            <Text style={[styles.activityDate, isRTL && styles.activityDateRTL]}>
                              {t.season}: {activity.season?.name || t.noSeasonAssigned}
                            </Text>
                            <TouchableOpacity
                              style={[
                                styles.activityStatusButton,
                                currentStatus === 'completed' && styles.activityStatusButtonCompleted
                              ]}
                              onPress={() => toggleActivityStatus(originalIndex)}
                            >
                              <Text style={[
                                styles.activityStatusText,
                                currentStatus === 'completed' && styles.activityStatusTextCompleted
                              ]}>
                                {currentStatus === 'completed' ? t.completed : t.pending}
                              </Text>
                            </TouchableOpacity>
                          </View>
                          <TouchableOpacity 
                            style={styles.removeActivityButton}
                            onPress={() => removeActivity(originalIndex)}
                          >
                            <Text style={styles.removeActivityButtonText}>✕</Text>
                          </TouchableOpacity>
                        </View>
                      );
                    })
                ) : (
                  <Text style={styles.hintText}>{t.noActivities}</Text>
                )}
              </View>
            </ScrollView>

            <View style={[styles.modalButtons, isRTL && styles.modalButtonsRTL]}>
              <TouchableOpacity
                style={[styles.modalButton, styles.cancelButton]}
                onPress={closeActivitiesModal}
              >
                <Text style={styles.cancelButtonText}>{t.close}</Text>
              </TouchableOpacity>
            </View>
          </View>
        </View>
      </Modal>
    </View>
  );
};

const styles = StyleSheet.create({
  container: {
    flex: 1,
    backgroundColor: '#f5f5f5',
    paddingTop: 40,
  },
  containerRTL: {
    direction: 'rtl',
  },
  tableRTL: {
    direction: 'rtl',
  },
  header: {
    backgroundColor: '#fff',
    padding: 20,
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    borderBottomWidth: 1,
    borderBottomColor: '#e0e0e0',
    shadowColor: '#000',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.1,
    shadowRadius: 4,
    elevation: 3,
  },
  headerRTL: {
    flexDirection: 'row-reverse',
  },
  headerRight: {
    flexDirection: 'row',
    alignItems: 'center',
  },
  title: {
    fontSize: 24,
    fontWeight: 'bold',
    color: '#333',
  },
  languageButton: {
    backgroundColor: '#2196F3',
    paddingHorizontal: 15,
    paddingVertical: 8,
    borderRadius: 5,
    marginRight: 10,
  },
  languageButtonText: {
    color: '#fff',
    fontSize: 14,
    fontWeight: '600',
  },
  addButton: {
    backgroundColor: '#4CAF50',
    paddingHorizontal: 20,
    paddingVertical: 10,
    borderRadius: 5,
  },
  addButtonText: {
    color: '#fff',
    fontSize: 16,
    fontWeight: '600',
  },
  scrollView: {
    flex: 1,
  },
  table: {
    backgroundColor: '#fff',
    margin: 20,
    borderRadius: 8,
    overflow: 'visible',
    shadowColor: '#000',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.1,
    shadowRadius: 4,
    elevation: 3,
  },
  tableHeader: {
    flexDirection: 'row',
    backgroundColor: '#f8f9fa',
    borderBottomWidth: 2,
    borderBottomColor: '#dee2e6',
    paddingVertical: 12,
    paddingHorizontal: 8,
  },
  tableHeaderRTL: {
    flexDirection: 'row-reverse',
  },
  seasonSectionHeader: {
    backgroundColor: '#e3f2fd',
    paddingVertical: 10,
    paddingHorizontal: 12,
    borderBottomWidth: 1,
    borderBottomColor: '#bbdefb',
  },
  seasonSectionHeaderRTL: {
    alignItems: 'flex-end',
  },
  seasonSectionHeaderText: {
    fontSize: 15,
    fontWeight: '700',
    color: '#1565c0',
  },
  seasonSectionHeaderTextRTL: {
    textAlign: 'right',
    width: '100%',
  },
  tableRow: {
    flexDirection: 'row',
    borderBottomWidth: 1,
    borderBottomColor: '#e9ecef',
    paddingVertical: 12,
    paddingHorizontal: 8,
    overflow: 'visible',
    position: 'relative',
    zIndex: 1,
  },
  tableRowRTL: {
    flexDirection: 'row-reverse',
  },
  statusButtonPending: {
    backgroundColor: '#FF9800',
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 4,
    alignSelf: 'flex-start',
  },
  statusButtonText: {
    color: '#fff',
    fontSize: 12,
    fontWeight: '600',
  },
  statusCompleted: {
    color: '#4CAF50',
    fontWeight: '600',
  },
  tableCell: {
    flex: 1,
    fontSize: 14,
    color: '#333',
    paddingHorizontal: 4,
    overflow: 'visible',
    position: 'relative',
    zIndex: 1,
  },
  tableCellText: {
    fontSize: 14,
    color: '#333',
  },
  tableCellRTL: {
    textAlign: 'right',
  },
  headerCell: {
    fontWeight: 'bold',
    color: '#495057',
  },
  headerCellText: {
    fontWeight: 'bold',
    color: '#495057',
  },
  headerCellRTL: {
    textAlign: 'right',
  },
  editButtonText: {
    color: '#fff',
    fontSize: 12,
    fontWeight: '600',
  },
  deleteButtonText: {
    color: '#fff',
    fontSize: 12,
    fontWeight: '600',
  },
  emptyState: {
    padding: 40,
    alignItems: 'center',
  },
  emptyText: {
    fontSize: 16,
    color: '#999',
    textAlign: 'center',
  },
  modalOverlay: {
    flex: 1,
    backgroundColor: 'rgba(0, 0, 0, 0.5)',
    justifyContent: 'center',
    alignItems: 'center',
    padding: 20,
  },
  modalContent: {
    backgroundColor: '#fff',
    borderRadius: 10,
    padding: 20,
    width: '100%',
    maxWidth: 800,
    maxHeight: '80%',
    minHeight: '40%',
  },
  modalTitle: {
    fontSize: 20,
    fontWeight: 'bold',
    marginBottom: 20,
    color: '#333',
  },
  form: {
    maxHeight: 400,
  },
  /** Bulk add activity modal: allow dropdown to extend and stay visible */
  bulkAddForm: {
    overflow: 'visible',
    maxHeight: 'none',
  },
  formGroup: {
    marginBottom: 16,
  },
  label: {
    fontSize: 14,
    fontWeight: '600',
    marginBottom: 8,
    color: '#333',
  },
  input: {
    borderWidth: 1,
    borderColor: '#ddd',
    borderRadius: 5,
    padding: 12,
    fontSize: 16,
    backgroundColor: '#fff',
  },
  textArea: {
    minHeight: 80,
    textAlignVertical: 'top',
  },
  radioGroup: {
    flexDirection: 'row',
    flexWrap: 'wrap',
  },
  radioOption: {
    paddingHorizontal: 16,
    paddingVertical: 8,
    borderRadius: 5,
    borderWidth: 1,
    borderColor: '#ddd',
    backgroundColor: '#fff',
    marginRight: 10,
    marginBottom: 10,
  },
  radioOptionSelected: {
    backgroundColor: '#2196F3',
    borderColor: '#2196F3',
  },
  radioText: {
    fontSize: 14,
    color: '#333',
  },
  radioTextSelected: {
    color: '#fff',
    fontWeight: '600',
  },
  switchGroup: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  switchGroupRTL: {
    flexDirection: 'row-reverse',
  },
  actionsCell: {
    flex: 1.5,
    flexDirection: 'row',
  },
  editButton: {
    backgroundColor: '#2196F3',
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 4,
    marginRight: 8,
  },
  deleteButton: {
    backgroundColor: '#f44336',
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 4,
  },
  modalButtons: {
    flexDirection: 'row',
    justifyContent: 'flex-end',
    marginTop: 20,
  },
  modalButtonsRTL: {
    flexDirection: 'row-reverse',
    justifyContent: 'flex-start',
  },
  modalContentRTL: {
    direction: 'rtl',
  },
  modalButton: {
    paddingHorizontal: 24,
    paddingVertical: 12,
    borderRadius: 5,
    marginLeft: 10,
  },
  cancelButton: {
    backgroundColor: '#e0e0e0',
  },
  cancelButtonText: {
    color: '#333',
    fontSize: 16,
    fontWeight: '600',
  },
  saveButton: {
    backgroundColor: '#4CAF50',
  },
  saveButtonText: {
    color: '#fff',
    fontSize: 16,
    fontWeight: '600',
  },
  kidsDetailsRow: {
    backgroundColor: '#f8f9fa',
    padding: 12,
    borderBottomWidth: 1,
    borderBottomColor: '#e9ecef',
    flexDirection: 'row',
    flexWrap: 'wrap',
    gap: 16,
  },
  kidsDetailsRowRTL: {
    direction: 'rtl',
  },
  kidsLabel: {
    fontSize: 12,
    fontWeight: 'bold',
    color: '#495057',
    marginBottom: 8,
  },
  kidsLabelRTL: {
    textAlign: 'right',
  },
  detailSection: {
    minWidth: 200,
    flex: 1,
  },
  kidItem: {
    marginBottom: 4,
  },
  kidText: {
    fontSize: 12,
    color: '#666',
    marginLeft: 12,
  },
  kidTextRTL: {
    marginLeft: 0,
    marginRight: 12,
    textAlign: 'right',
  },
  kidsHeader: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: 12,
  },
  kidsHeaderRTL: {
    flexDirection: 'row-reverse',
  },
  kidFormHeaderRTL: {
    flexDirection: 'row-reverse',
  },
  addKidButton: {
    backgroundColor: '#2196F3',
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 4,
  },
  addKidButtonText: {
    color: '#fff',
    fontSize: 12,
    fontWeight: '600',
  },
  birthdayRow: {
    flexDirection: 'row',
    gap: 8,
    marginTop: 4,
  },
  birthdayInput: {
    flex: 1,
    borderWidth: 1,
    borderColor: '#dee2e6',
    borderRadius: 4,
    paddingHorizontal: 10,
    paddingVertical: 8,
    fontSize: 14,
    minWidth: 0,
  },
  kidFormItem: {
    backgroundColor: '#f8f9fa',
    padding: 12,
    borderRadius: 5,
    marginBottom: 12,
    borderWidth: 1,
    borderColor: '#dee2e6',
  },
  kidFormHeader: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: 8,
  },
  kidFormLabel: {
    fontSize: 14,
    fontWeight: '600',
    color: '#333',
  },
  removeKidButton: {
    backgroundColor: '#f44336',
    paddingHorizontal: 10,
    paddingVertical: 4,
    borderRadius: 4,
  },
  removeKidButtonText: {
    color: '#fff',
    fontSize: 11,
    fontWeight: '600',
  },
  ageText: {
    fontSize: 12,
    color: '#4CAF50',
    marginTop: 4,
    fontStyle: 'italic',
  },
  hintText: {
    fontSize: 12,
    color: '#999',
    fontStyle: 'italic',
    textAlign: 'center',
    padding: 8,
  },
  kidsCountText: {
    fontSize: 12,
    color: '#666',
    marginTop: 8,
    fontWeight: '600',
  },
  expandButton: {
    padding: 4,
    marginRight: 8,
    minWidth: 24,
    alignItems: 'center',
    justifyContent: 'center',
  },
  expandIcon: {
    fontSize: 12,
    color: '#2196F3',
    fontWeight: 'bold',
  },
  searchContainer: {
    backgroundColor: '#fff',
    padding: 15,
    borderBottomWidth: 1,
    borderBottomColor: '#e0e0e0',
    flexDirection: 'row',
    alignItems: 'center',
  },
  searchContainerRTL: {
    flexDirection: 'row-reverse',
  },
  searchInput: {
    flex: 1,
    borderWidth: 1,
    borderColor: '#ddd',
    borderRadius: 8,
    padding: 12,
    fontSize: 16,
    backgroundColor: '#f8f9fa',
    paddingRight: 40,
  },
  searchInputRTL: {
    paddingRight: 12,
    paddingLeft: 40,
  },
  clearSearchButton: {
    position: 'absolute',
    right: 25,
    padding: 8,
    backgroundColor: '#e0e0e0',
    borderRadius: 12,
    width: 24,
    height: 24,
    justifyContent: 'center',
    alignItems: 'center',
  },
  clearSearchButtonText: {
    color: '#666',
    fontSize: 12,
    fontWeight: 'bold',
  },
  filtersContainer: {
    backgroundColor: '#fff',
    borderBottomWidth: 1,
    borderBottomColor: '#e0e0e0',
    position: 'relative',
    zIndex: 1,
  },
  filtersHeader: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: 15,
    backgroundColor: '#f8f9fa',
  },
  filtersHeaderRTL: {
    flexDirection: 'row-reverse',
  },
  filtersHeaderText: {
    fontSize: 16,
    fontWeight: '600',
    color: '#333',
  },
  filtersHeaderTextRTL: {
    textAlign: 'right',
  },
  clearFiltersButton: {
    paddingHorizontal: 12,
    paddingVertical: 6,
    backgroundColor: '#e0e0e0',
    borderRadius: 4,
  },
  clearFiltersButtonText: {
    color: '#666',
    fontSize: 12,
    fontWeight: '600',
  },
  filtersContent: {
    padding: 15,
    position: 'relative',
    zIndex: 1,
    overflow: 'visible',
  },
  filtersContentRTL: {
    direction: 'rtl',
  },
  filterRow: {
    flexDirection: 'row',
    marginBottom: 15,
    gap: 15,
    position: 'relative',
    zIndex: 1,
  },
  filterRowRTL: {
    flexDirection: 'row-reverse',
  },
  filterGroup: {
    flex: 1,
    position: 'relative',
    zIndex: 1,
    overflow: 'visible',
  },
  timelineFilterGroup: {
    minWidth: 200,
  },
  filterLabel: {
    fontSize: 14,
    fontWeight: '600',
    marginBottom: 8,
    color: '#333',
  },
  filterLabelRTL: {
    textAlign: 'right',
  },
  filterInput: {
    borderWidth: 1,
    borderColor: '#ddd',
    borderRadius: 5,
    padding: 10,
    fontSize: 14,
    backgroundColor: '#fff',
    zIndex: 1,
  },
  filterInputRTL: {
    textAlign: 'right',
  },
  inputRTL: {
    textAlign: 'right',
  },
  filterRadioGroup: {
    flexDirection: 'row',
    gap: 10,
  },
  filterRadioGroupRTL: {
    flexDirection: 'row-reverse',
  },
  filterRadioOption: {
    paddingHorizontal: 12,
    paddingVertical: 8,
    borderRadius: 5,
    borderWidth: 1,
    borderColor: '#ddd',
    backgroundColor: '#fff',
  },
  filterRadioOptionSelected: {
    backgroundColor: '#2196F3',
    borderColor: '#2196F3',
  },
  filterRadioText: {
    fontSize: 14,
    color: '#333',
  },
  filterRadioTextSelected: {
    color: '#fff',
    fontWeight: '600',
  },
  ageRangeContainer: {
    flexDirection: 'row',
    alignItems: 'center',
  },
  ageRangeContainerRTL: {
    flexDirection: 'row-reverse',
  },
  ageRangeSeparator: {
    fontSize: 14,
    color: '#666',
    marginHorizontal: 4,
  },
  paginationContainer: {
    backgroundColor: '#fff',
    padding: 15,
    borderTopWidth: 1,
    borderTopColor: '#e0e0e0',
    alignItems: 'center',
    justifyContent: 'center',
  },
  viewOptionsContainer: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 8,
    marginBottom: 15,
    justifyContent: 'center',
  },
  viewOptionsLabel: {
    fontSize: 14,
    fontWeight: '600',
    color: '#333',
    marginRight: 8,
  },
  viewOptionButton: {
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 4,
    borderWidth: 1,
    borderColor: '#ddd',
    backgroundColor: '#fff',
  },
  viewOptionButtonSelected: {
    backgroundColor: '#2196F3',
    borderColor: '#2196F3',
  },
  viewOptionText: {
    fontSize: 14,
    color: '#333',
  },
  viewOptionTextSelected: {
    color: '#fff',
    fontWeight: '600',
  },
  favoritesFilterContainer: {
    flexDirection: 'row',
    justifyContent: 'flex-start',
    alignItems: 'center',
    marginBottom: 12,
    gap: 8,
  },
  favoritesFilterContainerRTL: {
    flexDirection: 'row-reverse',
    justifyContent: 'flex-end',
  },
  favoritesFilterButton: {
    paddingHorizontal: 14,
    paddingVertical: 6,
    borderRadius: 20,
    borderWidth: 1,
    borderColor: '#2196F3',
    backgroundColor: '#fff',
    minWidth: 110,
    alignItems: 'center',
  },
  favoritesFilterButtonActive: {
    backgroundColor: '#2196F3',
  },
  favoritesFilterButtonText: {
    fontSize: 13,
    color: '#2196F3',
    fontWeight: '500',
  },
  favoritesFilterButtonTextActive: {
    color: '#fff',
    fontWeight: '600',
  },
  favoriteHeaderIcon: {
    fontSize: 16,
    color: '#FFC107',
  },
  favoriteIcon: {
    fontSize: 18,
    color: '#999',
  },
  favoriteIconActive: {
    color: '#FFC107',
  },
  paginationControls: {
    flexDirection: 'column',
    alignItems: 'center',
    gap: 10,
    width: '100%',
  },
  paginationInfo: {
    fontSize: 14,
    color: '#666',
    textAlign: 'center',
  },
  paginationButtons: {
    flexDirection: 'row',
    alignItems: 'center',
    gap: 12,
  },
  paginationButton: {
    paddingHorizontal: 16,
    paddingVertical: 8,
    borderRadius: 4,
    backgroundColor: '#2196F3',
  },
  paginationButtonDisabled: {
    backgroundColor: '#e0e0e0',
  },
  paginationButtonText: {
    color: '#fff',
    fontSize: 14,
    fontWeight: '600',
  },
  paginationButtonTextDisabled: {
    color: '#999',
  },
  pageNumber: {
    fontSize: 14,
    color: '#333',
    fontWeight: '600',
  },
  activitiesButton: {
    backgroundColor: '#FF9800',
    paddingHorizontal: 10,
    paddingVertical: 6,
    borderRadius: 4,
    marginRight: 8,
  },
  activitiesButtonText: {
    color: '#fff',
    fontSize: 14,
  },
  activityInputContainer: {
    flexDirection: 'row',
    alignItems: 'center',
  },
  activityInputContainerRTL: {
    flexDirection: 'row-reverse',
  },
  activityDescriptionRTL: {
    textAlign: 'right',
  },
  activityDateRTL: {
    textAlign: 'right',
  },
  activityStatusButton: {
    marginTop: 8,
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 5,
    backgroundColor: '#ff9800',
    alignSelf: 'flex-start',
  },
  activityStatusButtonCompleted: {
    backgroundColor: '#4CAF50',
  },
  activityStatusText: {
    color: '#fff',
    fontSize: 12,
    fontWeight: '600',
  },
  activityStatusTextCompleted: {
    color: '#fff',
  },
  bulkActivityButton: {
    backgroundColor: '#FF9800',
    paddingHorizontal: 15,
    paddingVertical: 8,
    borderRadius: 5,
    marginRight: 10,
  },
  bulkActivityButtonText: {
    color: '#fff',
    fontSize: 14,
    fontWeight: '600',
  },
  exportSelectedExcelButton: {
    backgroundColor: '#4CAF50',
    paddingHorizontal: 15,
    paddingVertical: 8,
    borderRadius: 5,
    marginRight: 10,
  },
  exportSelectedExcelButtonText: {
    color: '#fff',
    fontSize: 14,
    fontWeight: '600',
  },
  checkbox: {
    width: 20,
    height: 20,
    borderWidth: 2,
    borderColor: '#666',
    borderRadius: 3,
    backgroundColor: '#fff',
    justifyContent: 'center',
    alignItems: 'center',
  },
  checkboxChecked: {
    backgroundColor: '#2196F3',
    borderColor: '#2196F3',
  },
  checkboxCheckmark: {
    color: '#fff',
    fontSize: 14,
    fontWeight: 'bold',
  },
  addActivityButton: {
    backgroundColor: '#4CAF50',
    paddingHorizontal: 16,
    paddingVertical: 12,
    borderRadius: 5,
  },
  addActivityButtonText: {
    color: '#fff',
    fontSize: 14,
    fontWeight: '600',
  },
  activityItem: {
    backgroundColor: '#f8f9fa',
    padding: 12,
    borderRadius: 5,
    marginBottom: 8,
    borderWidth: 1,
    borderColor: '#dee2e6',
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  activityContent: {
    flex: 1,
    marginRight: 8,
  },
  activityDescription: {
    fontSize: 14,
    color: '#333',
    marginBottom: 4,
  },
  activityDate: {
    fontSize: 12,
    color: '#666',
    fontStyle: 'italic',
  },
  removeActivityButton: {
    backgroundColor: '#f44336',
    width: 28,
    height: 28,
    borderRadius: 14,
    justifyContent: 'center',
    alignItems: 'center',
  },
  removeActivityButtonText: {
    color: '#fff',
    fontSize: 16,
    fontWeight: 'bold',
  },
  navButton: {
    backgroundColor: '#9C27B0',
    paddingHorizontal: 15,
    paddingVertical: 8,
    borderRadius: 5,
    marginRight: 10,
  },
  navButtonText: {
    color: '#fff',
    fontSize: 14,
    fontWeight: '600',
  },
  exportExcelButton: {
    backgroundColor: '#4CAF50',
    paddingHorizontal: 15,
    paddingVertical: 8,
    borderRadius: 5,
    marginRight: 10,
  },
  exportExcelButtonText: {
    color: '#fff',
    fontSize: 14,
    fontWeight: '600',
  },
  manageDistributorsButton: {
    backgroundColor: '#FF5722',
    paddingHorizontal: 15,
    paddingVertical: 8,
    borderRadius: 5,
    marginRight: 10,
  },
  manageDistributorsButtonText: {
    color: '#fff',
    fontSize: 14,
    fontWeight: '600',
  },
  distributorSelector: {
    width: '100%',
    position: 'relative',
    zIndex: 9999999,
  },
  distributorSelectButton: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    width: '100%',
    padding: 8,
    borderWidth: 1,
    borderColor: '#ddd',
    borderRadius: 5,
    backgroundColor: '#fff',
    minHeight: 36,
  },
  distributorSelectButtonText: {
    fontSize: 14,
    color: '#333',
    flex: 1,
  },
  distributorSelectArrow: {
    fontSize: 10,
    color: '#666',
    marginLeft: 8,
  },
  distributorDropdown: {
    position: 'absolute',
    top: '100%',
    left: 0,
    right: 0,
    backgroundColor: '#fff',
    borderWidth: 1,
    borderColor: '#ddd',
    borderRadius: 5,
    marginTop: 4,
    maxHeight: 300,
    zIndex: 9999999,
    shadowColor: '#000',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.3,
    shadowRadius: 10,
    elevation: 9999,
    overflow: 'hidden',
  },
  /** Wrapper for distributor in bulk-add modal so dropdown is not clipped and stacks above buttons */
  bulkAddDistributorFormGroup: {
    position: 'relative',
    zIndex: 10000,
    overflow: 'visible',
  },
  distributorDropdownScroll: {
    maxHeight: 300,
  },
  distributorOption: {
    padding: 12,
    borderBottomWidth: 1,
    borderBottomColor: '#f0f0f0',
  },
  distributorOptionText: {
    fontSize: 14,
    color: '#333',
  },
  distributorItem: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    backgroundColor: '#f8f9fa',
    padding: 12,
    borderRadius: 5,
    marginBottom: 8,
    borderWidth: 1,
    borderColor: '#dee2e6',
  },
  distributorItemInfo: {
    flex: 1,
  },
  distributorItemName: {
    fontSize: 16,
    fontWeight: '600',
    color: '#333',
    marginBottom: 4,
  },
  distributorItemPhone: {
    fontSize: 14,
    color: '#666',
  },
  removeDistributorButton: {
    backgroundColor: '#f44336',
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 4,
  },
  removeDistributorButtonText: {
    color: '#fff',
    fontSize: 12,
    fontWeight: '600',
  },
  addDistributorButton: {
    backgroundColor: '#4CAF50',
    paddingHorizontal: 16,
    paddingVertical: 10,
    borderRadius: 5,
    marginTop: 8,
    alignItems: 'center',
  },
  addDistributorButtonText: {
    color: '#fff',
    fontSize: 14,
    fontWeight: '600',
  },
});

export default App;
