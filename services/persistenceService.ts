import { CriterionKey, Student, Unit } from '../types';
import { createEmptyTeacherObservations } from './teacherObservationScales';

const STORAGE_KEY = 'termgenius-session';
const VERSION = 1;

interface SerializedFile {
  name: string;
  mimeType: string;
  dataBase64: string;
}

interface SerializedCriterionConfig {
  enabled: boolean;
  notes: string;
  file: SerializedFile | null;
}

interface SerializedUnit {
  id: string;
  title: string;
  criteria: Record<CriterionKey, SerializedCriterionConfig>;
}

export interface PersistedSession {
  version: number;
  students: Student[];
  units: SerializedUnit[];
  activeFile: string | null;
  showConfig: boolean;
}

export const createDefaultUnits = (): Unit[] => [{
  id: 'default-1',
  title: '',
  criteria: {
    A: { enabled: true, file: null, notes: '' },
    B: { enabled: true, file: null, notes: '' },
    C: { enabled: true, file: null, notes: '' },
    D: { enabled: true, file: null, notes: '' },
  },
}];

const isDefaultUnits = (units: Unit[]): boolean => {
  if (units.length !== 1) return false;
  const unit = units[0];
  if (unit.title.trim()) return false;
  return (['A', 'B', 'C', 'D'] as CriterionKey[]).every((key) => {
    const crit = unit.criteria[key];
    return crit.enabled && !crit.notes.trim() && !crit.file;
  });
};

export const hasMeaningfulData = (session: {
  students: Student[];
  units: Unit[];
  activeFile: string | null;
}): boolean => {
  if (session.students.length > 0) return true;
  if (session.activeFile) return true;
  return !isDefaultUnits(session.units);
};

const fileToBase64 = (file: File): Promise<string> =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      resolve(result.split(',')[1]);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });

const base64ToFile = ({ name, mimeType, dataBase64 }: SerializedFile): File => {
  const binary = atob(dataBase64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return new File([bytes], name, { type: mimeType });
};

const serializeUnits = async (units: Unit[]): Promise<SerializedUnit[]> =>
  Promise.all(
    units.map(async (unit) => ({
      id: unit.id,
      title: unit.title,
      criteria: await Promise.all(
        (['A', 'B', 'C', 'D'] as CriterionKey[]).map(async (key) => {
          const crit = unit.criteria[key];
          return [
            key,
            {
              enabled: crit.enabled,
              notes: crit.notes,
              file: crit.file
                ? {
                    name: crit.file.name,
                    mimeType: crit.file.type || 'application/octet-stream',
                    dataBase64: await fileToBase64(crit.file),
                  }
                : null,
            },
          ] as const;
        })
      ).then((entries) => Object.fromEntries(entries) as Record<CriterionKey, SerializedCriterionConfig>),
    }))
  );

const deserializeUnits = (units: SerializedUnit[]): Unit[] =>
  units.map((unit) => ({
    id: unit.id,
    title: unit.title,
    criteria: Object.fromEntries(
      (['A', 'B', 'C', 'D'] as CriterionKey[]).map((key) => {
        const crit = unit.criteria[key];
        return [
          key,
          {
            enabled: crit.enabled,
            notes: crit.notes,
            file: crit.file ? base64ToFile(crit.file) : null,
          },
        ];
      })
    ) as Record<CriterionKey, Unit['criteria'][CriterionKey]>,
  }));

const normalizeStudent = (student: Student): Student => {
  const normalized: Student = {
    ...student,
    teacherObservations: {
      ...createEmptyTeacherObservations(),
      ...student.teacherObservations,
    },
  };

  if (normalized.status !== 'generating') return normalized;
  return {
    ...normalized,
    status: normalized.generatedSummary?.trim() ? 'completed' : 'idle',
  };
};

export const loadSession = (): {
  students: Student[];
  units: Unit[];
  activeFile: string | null;
  showConfig: boolean;
} | null => {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;

    const data = JSON.parse(raw) as PersistedSession;
    if (data.version !== VERSION) return null;

    return {
      students: (data.students ?? []).map(normalizeStudent),
      units: data.units?.length ? deserializeUnits(data.units) : createDefaultUnits(),
      activeFile: data.activeFile ?? null,
      showConfig: data.showConfig ?? true,
    };
  } catch (error) {
    console.warn('Failed to load saved session', error);
    return null;
  }
};

export const saveSession = async (session: {
  students: Student[];
  units: Unit[];
  activeFile: string | null;
  showConfig: boolean;
}): Promise<void> => {
  if (!hasMeaningfulData(session)) {
    clearSession();
    return;
  }

  const payload: PersistedSession = {
    version: VERSION,
    students: session.students.map(normalizeStudent),
    units: await serializeUnits(session.units),
    activeFile: session.activeFile,
    showConfig: session.showConfig,
  };

  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  } catch (error) {
    console.warn('Failed to save session (storage may be full)', error);
  }
};

export const clearSession = (): void => {
  localStorage.removeItem(STORAGE_KEY);
};

export const hasSavedSession = (): boolean => {
  try {
    return localStorage.getItem(STORAGE_KEY) !== null;
  } catch {
    return false;
  }
};
