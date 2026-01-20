import React from 'react';
import { Unit, CriterionKey, CriterionConfig } from '../types';
import { Plus, Trash2, Upload, FileText, CheckSquare, Square, X } from 'lucide-react';

interface UnitConfigurationProps {
  units: Unit[];
  setUnits: React.Dispatch<React.SetStateAction<Unit[]>>;
}

export const UnitConfiguration: React.FC<UnitConfigurationProps> = ({ units, setUnits }) => {
  const addUnit = () => {
    const newUnit: Unit = {
      id: crypto.randomUUID(),
      title: '',
      criteria: {
        A: { enabled: true, file: null, notes: '' },
        B: { enabled: true, file: null, notes: '' },
        C: { enabled: true, file: null, notes: '' },
        D: { enabled: true, file: null, notes: '' },
      },
    };
    setUnits([...units, newUnit]);
  };

  const removeUnit = (id: string) => {
    setUnits(units.filter((u) => u.id !== id));
  };

  const updateUnitTitle = (id: string, title: string) => {
    setUnits(units.map((u) => (u.id === id ? { ...u, title } : u)));
  };

  const updateCriterion = (
    unitId: string,
    key: CriterionKey,
    updates: Partial<CriterionConfig>
  ) => {
    setUnits(
      units.map((u) => {
        if (u.id !== unitId) return u;
        return {
          ...u,
          criteria: {
            ...u.criteria,
            [key]: { ...u.criteria[key], ...updates },
          },
        };
      })
    );
  };

  return (
    <div className="space-y-8 w-full mb-8">
      <div className="flex justify-between items-center mb-4">
        <h3 className="text-xl font-bold text-slate-800">Unit & Criterion Configuration</h3>
      </div>

      {units.map((unit, index) => (
        <div key={unit.id} className="bg-white rounded-xl border border-slate-200 shadow-sm p-6 relative">
          {units.length > 1 && (
            <button
              onClick={() => removeUnit(unit.id)}
              className="absolute top-4 right-4 text-slate-400 hover:text-red-500 transition-colors"
              title="Remove Unit"
            >
              <Trash2 className="w-5 h-5" />
            </button>
          )}

          <div className="mb-6">
            <label className="block text-sm font-medium text-slate-700 mb-1">
              Unit Title
            </label>
            <input
              type="text"
              value={unit.title}
              onChange={(e) => updateUnitTitle(unit.id, e.target.value)}
              placeholder="e.g., Ancient Civilizations, Thermodynamics..."
              className="w-full max-w-md px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
            />
          </div>

          <div className="grid grid-cols-2 gap-4">
            {(['A', 'B', 'C', 'D'] as CriterionKey[]).map((key) => {
              const config = unit.criteria[key];
              const isEnabled = config.enabled;

              return (
                <div
                  key={key}
                  className={`border rounded-lg p-4 flex flex-col gap-3 transition-colors ${
                    isEnabled ? 'border-slate-200 bg-slate-50' : 'border-slate-200 bg-slate-100'
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className={`font-bold ${isEnabled ? 'text-slate-700' : 'text-slate-400'}`}>Criterion {key}</span>
                  </div>

                  {/* Upload Box */}
                  <div className="relative">
                    <input
                      type="file"
                      disabled={!isEnabled}
                      onChange={(e) =>
                        updateCriterion(unit.id, key, {
                          file: e.target.files ? e.target.files[0] : null,
                        })
                      }
                      className="hidden"
                      id={`file-${unit.id}-${key}`}
                    />
                    <label
                      htmlFor={isEnabled && !config.file ? `file-${unit.id}-${key}` : undefined}
                      className={`flex flex-col items-center justify-center h-24 border-2 border-dashed rounded-lg transition-colors ${
                        isEnabled
                          ? config.file 
                            ? 'border-blue-300 bg-blue-50'
                            : 'border-slate-300 hover:border-blue-500 hover:bg-blue-50 cursor-pointer bg-white'
                          : 'border-slate-300 bg-slate-200 cursor-not-allowed'
                      }`}
                    >
                      {config.file ? (
                        <div className="text-center px-2 relative w-full">
                          <FileText className={`w-6 h-6 mx-auto mb-1 ${isEnabled ? 'text-blue-600' : 'text-slate-400'}`} />
                          <p className={`text-xs truncate max-w-[120px] mx-auto ${isEnabled ? 'text-slate-600' : 'text-slate-400'}`}>
                            {config.file.name}
                          </p>
                          {isEnabled && (
                            <button
                              type="button"
                              onClick={(e) => {
                                e.preventDefault();
                                e.stopPropagation();
                                updateCriterion(unit.id, key, { file: null });
                              }}
                              className="absolute -top-1 -right-1 p-1 bg-red-100 hover:bg-red-200 rounded-full text-red-600 hover:text-red-700 transition-colors"
                              title="Remove file"
                            >
                              <X className="w-4 h-4" />
                            </button>
                          )}
                        </div>
                      ) : (
                        <div className="text-center">
                          <Upload className={`w-6 h-6 mx-auto mb-1 ${isEnabled ? 'text-slate-400' : 'text-slate-300'}`} />
                          <p className={`text-xs ${isEnabled ? 'text-slate-500' : 'text-slate-400'}`}>Task Clarification</p>
                        </div>
                      )}
                    </label>
                  </div>

                  {/* Notes Section */}
                  <div>
                    <label className={`block text-xs font-semibold mb-1 ${isEnabled ? 'text-slate-500' : 'text-slate-400'}`}>
                      Teacher Notes
                    </label>
                    <textarea
                      disabled={!isEnabled}
                      value={config.notes}
                      onChange={(e) =>
                        updateCriterion(unit.id, key, { notes: e.target.value })
                      }
                      placeholder={isEnabled ? "Specific focus areas..." : "Disabled"}
                      className={`w-full text-sm px-2 py-1.5 border rounded-md focus:ring-1 focus:ring-blue-500 focus:outline-none min-h-[60px] ${
                          isEnabled ? 'border-slate-300 bg-white' : 'border-slate-200 bg-slate-100 text-slate-400 resize-none'
                      }`}
                    />
                  </div>

                  {/* Toggle Checkbox */}
                  <div 
                    onClick={() => updateCriterion(unit.id, key, { enabled: !isEnabled })}
                    className="flex items-center gap-2 cursor-pointer mt-auto pt-2 border-t border-slate-200 select-none group"
                  >
                    <div className="text-slate-500 group-hover:text-slate-700">
                        {!isEnabled ? <CheckSquare className="w-5 h-5 text-slate-600" /> : <Square className="w-5 h-5 text-slate-400" />}
                    </div>
                    <span className="text-xs text-slate-500">
                       No criterion for this unit
                    </span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      ))}

      <button
        onClick={addUnit}
        className="w-full py-3 bg-white border-2 border-dashed border-slate-300 rounded-xl text-slate-500 font-medium hover:border-blue-500 hover:text-blue-600 hover:bg-blue-50 transition-all flex items-center justify-center gap-2"
      >
        <Plus className="w-5 h-5" />
        Add another Unit
      </button>
    </div>
  );
};