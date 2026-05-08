import { describe, it, expect, vi } from 'vitest';
import MicrosoftConnect from '../src/MicrosoftConnect';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import api from '../src/api';

vi.mock('../src/api');

describe('MicrosoftConnect Component', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('renders connect button when disconnected', async () => {
    api.getStatus.mockResolvedValue({ connected: false });

    render(<MicrosoftConnect />);

    await waitFor(() => {
      expect(screen.getByText('Connect Microsoft Account')).toBeInTheDocument();
    });
  });

  it('renders connected status when connected', async () => {
    api.getStatus.mockResolvedValue({
      connected: true,
      microsoft_user_id: 'user@example.com',
    });

    render(<MicrosoftConnect />);

    await waitFor(() => {
      expect(screen.getByText(/✓ Connected/)).toBeInTheDocument();
      expect(screen.getByText(/user@example.com/)).toBeInTheDocument();
    });
  });

  it('calls connect endpoint on button click', async () => {
    api.getStatus.mockResolvedValue({ connected: false });
    api.getConnectUrl.mockResolvedValue({
      auth_url: 'https://login.microsoftonline.com/...',
      state: 'state123',
    });

    const { rerender } = render(<MicrosoftConnect />);

    const button = await screen.findByText('Connect Microsoft Account');
    fireEvent.click(button);

    await waitFor(() => {
      expect(api.getConnectUrl).toHaveBeenCalled();
    });
  });

  it('renders error on connection failure', async () => {
    api.getStatus.mockResolvedValue({ connected: false });
    api.getConnectUrl.mockRejectedValue(new Error('Network error'));

    render(<MicrosoftConnect />);

    const button = await screen.findByText('Connect Microsoft Account');
    fireEvent.click(button);

    await waitFor(() => {
      expect(screen.getByText(/Failed to initiate connection/)).toBeInTheDocument();
    });
  });
});
