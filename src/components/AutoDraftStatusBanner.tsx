import React from 'react';

interface AutoDraftStatusBadgeProps {
  menuLabel?: string;
  lastDraftTimeStr: string;
  isDirty?: boolean;
  isOnline?: boolean;
  onManualSave?: () => void;
  onRestoreDraft?: () => void;
  onClearDraft?: () => void;
  hasRestorableDraft?: boolean;
  restorableDraftSummary?: string;
  variant?: 'light' | 'dark' | 'banner';
}

export const AutoDraftStatusBadge: React.FC<AutoDraftStatusBadgeProps> = () => {
  return null;
};

