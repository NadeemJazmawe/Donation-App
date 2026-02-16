import React, { useState, useEffect } from 'react';
import { View, Text, StyleSheet, ScrollView, TouchableOpacity, Modal, TextInput, Switch } from 'react-native';
import { translations } from './translations';
import * as XLSX from 'xlsx';

const API_URL = 'http://localhost:5000/api/persons';

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
    activityStatus: 'open', // 'all', 'open', 'closed'
    activityName: ''
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
  const [expandedKids, setExpandedKids] = useState({}); // Track which persons have kids expanded
  const [newActivity, setNewActivity] = useState('');
  const [formData, setFormData] = useState({
    name: '',
    phone: '',
    kids: [],
    monthIncome: '',
    maritalStatus: 'Single',
    liveInRenta: false,
    hasCar: false,
    bankNumber: ''
  });

  useEffect(() => {
    fetchPersons();
  }, []);

  // Save distributors to localStorage whenever they change
  useEffect(() => {
    localStorage.setItem('distributors', JSON.stringify(distributors));
  }, [distributors]);

  // Filter persons based on search query and filters
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

    setFilteredPersons(filtered);
    setCurrentPage(1); // Reset to first page when filters change
  }, [searchQuery, filters, persons]);

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
          peopleActivities.push({
            ...activity,
            personId: person._id,
            personName: person.name,
            personKidsNumber: person.kidsNumber || 0,
            activityIndex: activityIndex,
            activityStatus: activity.status || 'pending'
          });
        });
      } else {
        // Include people with no activities
        peopleActivities.push({
          personId: person._id,
          personName: person.name,
          personKidsNumber: person.kidsNumber || 0,
          description: '',
          distributor: '',
          activityStatus: 'none',
          date: person.createdAt || new Date().toISOString()
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

    return activities;
  };

  // Export filtered activities to Excel
  const exportToExcel = () => {
    const filteredActivities = getFilteredActivities();
    
    // Prepare data for Excel
    const excelData = filteredActivities.map(activity => ({
      [t.name]: activity.personName,
      [t.kids]: activity.personKidsNumber,
      [t.activity]: activity.description || '-',
      [t.distributor]: activity.distributor || '-',
      [t.status]: activity.activityStatus === 'pending' || !activity.activityStatus ? t.pending : 
                  activity.activityStatus === 'completed' ? t.completed : '-',
      [t.date]: activity.date ? new Date(activity.date).toLocaleString() : '-'
    }));

    // Create workbook and worksheet
    const worksheet = XLSX.utils.json_to_sheet(excelData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, t.openActivities);

    // Generate Excel file
    const fileName = `Open_Activities_${new Date().toISOString().split('T')[0]}.xlsx`;
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
        body: JSON.stringify({
          ...person,
          activities: updatedActivities
        }),
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
        body: JSON.stringify({
          ...person,
          activities: updatedActivities
        }),
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

  const handleAdd = () => {
    setEditingPerson(null);
    setFormData({
      name: '',
      phone: '',
      kids: [],
      monthIncome: '',
      maritalStatus: 'Single',
      liveInRenta: false,
      hasCar: false,
      bankNumber: ''
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
      monthIncome: person.monthIncome.toString(),
      maritalStatus: person.maritalStatus,
      liveInRenta: person.liveInRenta,
      hasCar: person.hasCar,
      bankNumber: person.bankNumber
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
        monthIncome: parseFloat(formData.monthIncome) || 0,
        maritalStatus: formData.maritalStatus,
        liveInRenta: formData.liveInRenta,
        hasCar: formData.hasCar,
        bankNumber: formData.bankNumber
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
        alert('Error: ' + (error.message || 'Failed to save'));
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
    setActivitiesModalVisible(true);
  };

  // Close activities modal
  const closeActivitiesModal = () => {
    setActivitiesModalVisible(false);
    setSelectedPersonForActivities(null);
    setNewActivity('');
  };

  // Add new activity
  const addActivity = async () => {
    if (!newActivity.trim() || !selectedPersonForActivities) return;

    try {
      const updatedActivities = [
        ...(selectedPersonForActivities.activities || []),
        {
          description: newActivity.trim(),
          date: new Date().toISOString(),
          status: 'pending'
        }
      ];

      const response = await fetch(`${API_URL}/${selectedPersonForActivities._id}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          ...selectedPersonForActivities,
          activities: updatedActivities
        }),
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

  // Add activity to selected persons
  const addActivityToSelected = async (activityDescription) => {
    if (!activityDescription.trim() || selectedPersonsForActivities.length === 0) {
      alert(t.selectPeopleFirst || 'Please select at least one person first');
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
            status: 'pending'
          }
        ];

        const response = await fetch(`${API_URL}/${personId}`, {
          method: 'PUT',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            ...person,
            activities: updatedActivities
          }),
        });

        if (!response.ok) {
          throw new Error(`Failed to update person ${person.name}`);
        }
      });

      await Promise.all(promises);
      setSelectedPersonsForActivities([]);
      fetchPersons(); // Refresh the list
      alert(t.activityAddedToSelected || `Activity added to ${selectedPersonsForActivities.length} person(s)`);
    } catch (error) {
      console.error('Error adding activity to selected persons:', error);
      alert(t.errorAddingActivity);
    }
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
        body: JSON.stringify({
          ...selectedPersonForActivities,
          activities: updatedActivities
        }),
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
        body: JSON.stringify({
          ...selectedPersonForActivities,
          activities: updatedActivities
        }),
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

  // Render Open Activities page
  const renderOpenActivitiesPage = () => {
    const filteredActivities = getFilteredActivities();

    return (
      <View style={[styles.container, isRTL && styles.containerRTL]}>
        <View style={[styles.header, isRTL && styles.headerRTL]}>
          <Text style={styles.title}>{t.openActivities}</Text>
          <View style={styles.headerRight}>
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
              style={styles.navButton}
              onPress={() => setCurrentView('main')}
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
                  activityStatus: 'open',
                  activityName: ''
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
              <View style={[styles.filterRow, isRTL && styles.filterRowRTL, { zIndex: 1 }]}>
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
            </View>
          )}
        </View>

        <ScrollView 
          style={styles.scrollView}
          onScrollBeginDrag={() => setOpenDropdowns({})}
        >
          <View style={[styles.table, isRTL && styles.tableRTL]}>
            <View style={[styles.tableHeader, isRTL && styles.tableHeaderRTL]}>
              <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL, { flex: 2 }]}>{t.name}</Text>
              <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL]}>{t.kids}</Text>
              <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL, { flex: 2 }]}>{t.activity}</Text>
              <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL, { flex: 2 }]}>{t.distributor}</Text>
              <Text style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL]}>{t.status}</Text>
            </View>

            {filteredActivities.length === 0 ? (
              <View style={styles.emptyState}>
                <Text style={styles.emptyText}>{t.noActivitiesFound}</Text>
              </View>
            ) : (
              filteredActivities.map((activity, index) => {
                const isPending = (activity.activityStatus === 'pending' || !activity.activityStatus) && activity.description;
                const isCompleted = activity.activityStatus === 'completed';
                const hasNoActivity = !activity.description || activity.activityStatus === 'none';
                
                const isDropdownOpen = openDropdowns[`${activity.personId}-${activity.activityIndex}`];
                return (
                  <View key={`${activity.personId}-${activity.date}-${index}`} style={[styles.tableRow, isRTL && styles.tableRowRTL, { zIndex: isDropdownOpen ? 99999999 : 1 }]}>
                    <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL, { flex: 2 }]}>
                      {activity.personName}
                    </Text>
                    <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>
                      {activity.personKidsNumber}
                    </Text>
                    <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL, { flex: 2 }]}>
                      {activity.description || '-'}
                    </Text>
                  <View style={[styles.tableCell, { flex: 2, zIndex: openDropdowns[`${activity.personId}-${activity.activityIndex}`] ? 99999999 : 1 }]}>
                    {activity.activityStatus !== 'none' ? (
                      <View style={[styles.distributorSelector, { zIndex: openDropdowns[`${activity.personId}-${activity.activityIndex}`] ? 99999999 : 1 }]}>
                        <TouchableOpacity
                          style={styles.distributorSelectButton}
                          onPress={(e) => {
                            e.stopPropagation();
                            const dropdownKey = `${activity.personId}-${activity.activityIndex}`;
                            setOpenDropdowns(prev => ({
                              ...prev,
                              [dropdownKey]: !prev[dropdownKey]
                            }));
                          }}
                        >
                          <Text style={styles.distributorSelectButtonText}>
                            {activity.distributor || t.selectDistributor}
                          </Text>
                          <Text style={styles.distributorSelectArrow}>
                            {openDropdowns[`${activity.personId}-${activity.activityIndex}`] ? '▲' : '▼'}
                          </Text>
                        </TouchableOpacity>
                        {openDropdowns[`${activity.personId}-${activity.activityIndex}`] && (
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
                                    [`${activity.personId}-${activity.activityIndex}`]: false
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
                                      [`${activity.personId}-${activity.activityIndex}`]: false
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
              })
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
            <TouchableOpacity 
              style={styles.bulkActivityButton}
              onPress={() => {
                const activityDescription = prompt(t.enterActivityDescription || 'Enter activity description:');
                if (activityDescription && activityDescription.trim()) {
                  addActivityToSelected(activityDescription);
                }
              }}
            >
              <Text style={styles.bulkActivityButtonText}>
                {t.addActivityToSelected || 'Add Activity'} ({selectedPersonsForActivities.length})
              </Text>
            </TouchableOpacity>
          )}
          <TouchableOpacity 
            style={styles.navButton}
            onPress={() => {
              fetchPersons(); // Refresh data before showing
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
            <View style={[styles.tableCell, styles.headerCell, isRTL && styles.headerCellRTL, { width: 50, flex: 0, minWidth: 50, maxWidth: 50 }]}>
              <Text style={[styles.headerCellText, isRTL && styles.headerCellRTL]}></Text>
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
            const isExpanded = expandedKids[person._id];
            const isSelected = selectedPersonsForActivities.includes(person._id);
            
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
                  <View style={[styles.tableCell, { flex: 2, flexDirection: isRTL ? 'row-reverse' : 'row', alignItems: 'center' }]}>
                    {hasKids && (
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
                  <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>{person.monthIncome}</Text>
                  <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>{t[person.maritalStatus.toLowerCase()] || person.maritalStatus}</Text>
                  <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>{person.liveInRenta ? t.yes : t.no}</Text>
                  <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>{person.hasCar ? t.yes : t.no}</Text>
                  <Text style={[styles.tableCell, styles.tableCellText, isRTL && styles.tableCellRTL]}>{person.bankNumber}</Text>
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
                {hasKids && isExpanded && (
                  <View style={[styles.kidsDetailsRow, isRTL && styles.kidsDetailsRowRTL]}>
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
                    <TextInput
                      style={[styles.input, isRTL && styles.inputRTL]}
                      value={kid.birthday}
                      onChangeText={(text) => updateKid(index, 'birthday', text)}
                      placeholder={t.birthday}
                      keyboardType="default"
                      textAlign={isRTL ? 'right' : 'left'}
                    />
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
                <View style={[styles.switchGroup, isRTL && styles.switchGroupRTL]}>
                  <Text style={styles.label}>{t.liveInRenta}</Text>
                  <Switch
                    value={formData.liveInRenta}
                    onValueChange={(value) => setFormData({ ...formData, liveInRenta: value })}
                  />
                </View>
              </View>

              <View style={styles.formGroup}>
                <View style={[styles.switchGroup, isRTL && styles.switchGroupRTL]}>
                  <Text style={styles.label}>{t.hasCar}</Text>
                  <Switch
                    value={formData.hasCar}
                    onValueChange={(value) => setFormData({ ...formData, hasCar: value })}
                  />
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
