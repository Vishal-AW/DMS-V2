import * as React from "react";
export type StatusType = 'draft' | 'pending' | 'approved' | 'rejected' | 'processing' | 'completed' | 'failed';

interface StatusBadgeProps {
  status: StatusType;
  label?: string;
}

const statusLabels: Record<StatusType, string> = {
  draft: 'Draft',
  pending: 'Pending Review',
  approved: 'Approved',
  rejected: 'Rejected',
  processing: 'Processing',
  completed: 'Completed',
  failed: 'Failed',
};

export default function StatusBadge({ status, label }: StatusBadgeProps) {
  return (
    <span
      className={`status-badge status-badge-${status}`}
      data-testid={`badge-status-${status}`}
    >
      {label || statusLabels[status]}
    </span>
  );
}
