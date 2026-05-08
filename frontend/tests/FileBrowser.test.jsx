import { describe, it, expect, vi } from 'vitest';
import FileBrowser from '../src/FileBrowser';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import api from '../src/api';

vi.mock('../src/api');

describe('FileBrowser Component', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('shows disconnected message when not connected', async () => {
    api.getStatus.mockResolvedValue({ connected: false });

    render(<FileBrowser />);

    await waitFor(() => {
      expect(
        screen.getByText(/Connect your Microsoft account/)
      ).toBeInTheDocument();
    });
  });

  it('loads sites on connect', async () => {
    api.getStatus.mockResolvedValue({ connected: true });
    api.listSites.mockResolvedValue({
      value: [
        { id: 'site1', displayName: 'Site 1' },
      ],
    });

    render(<FileBrowser />);

    await waitFor(() => {
      expect(screen.getByText('Site 1')).toBeInTheDocument();
    });
  });

  it('loads drives when site selected', async () => {
    api.getStatus.mockResolvedValue({ connected: true });
    api.listSites.mockResolvedValue({
      value: [
        { id: 'site1', displayName: 'Site 1' },
      ],
    });
    api.listDrives.mockResolvedValue({
      value: [
        { id: 'drive1', name: 'Documents' },
      ],
    });

    render(<FileBrowser />);

    const siteButton = await screen.findByText('Site 1');
    fireEvent.click(siteButton);

    await waitFor(() => {
      expect(screen.getByText('Documents')).toBeInTheDocument();
    });
  });

  it('loads items when drive selected', async () => {
    api.getStatus.mockResolvedValue({ connected: true });
    api.listSites.mockResolvedValue({
      value: [{ id: 'site1', displayName: 'Site 1' }],
    });
    api.listDrives.mockResolvedValue({
      value: [{ id: 'drive1', name: 'Documents' }],
    });
    api.listItems.mockResolvedValue({
      value: [
        { id: 'item1', name: 'folder', folder: {} },
        { id: 'file1', name: 'document.pdf' },
      ],
    });

    render(<FileBrowser />);

    const siteButton = await screen.findByText('Site 1');
    fireEvent.click(siteButton);

    const driveButton = await screen.findByText('Documents');
    fireEvent.click(driveButton);

    await waitFor(() => {
      expect(screen.getByText('document.pdf')).toBeInTheDocument();
    });
  });

  it('shows folder and file icons correctly', async () => {
    api.getStatus.mockResolvedValue({ connected: true });
    api.listSites.mockResolvedValue({
      value: [{ id: 'site1', displayName: 'Site 1' }],
    });
    api.listDrives.mockResolvedValue({
      value: [{ id: 'drive1', name: 'Documents' }],
    });
    api.listItems.mockResolvedValue({
      value: [
        { id: 'folder1', name: 'MyFolder', folder: {} },
        { id: 'file1', name: 'document.pdf' },
      ],
    });

    render(<FileBrowser />);

    const siteButton = await screen.findByText('Site 1');
    fireEvent.click(siteButton);

    const driveButton = await screen.findByText('Documents');
    fireEvent.click(driveButton);

    await waitFor(() => {
      expect(screen.getByText('📁')).toBeInTheDocument();
      expect(screen.getByText('📄')).toBeInTheDocument();
    });
  });

  it('calls download API when download button clicked', async () => {
    api.getStatus.mockResolvedValue({ connected: true });
    api.listSites.mockResolvedValue({
      value: [{ id: 'site1', displayName: 'Site 1' }],
    });
    api.listDrives.mockResolvedValue({
      value: [{ id: 'drive1', name: 'Documents' }],
    });
    api.listItems.mockResolvedValue({
      value: [{ id: 'file1', name: 'document.pdf' }],
    });
    api.downloadItem.mockResolvedValue(undefined);

    render(<FileBrowser />);

    const siteButton = await screen.findByText('Site 1');
    fireEvent.click(siteButton);

    const driveButton = await screen.findByText('Documents');
    fireEvent.click(driveButton);

    const downloadButton = await screen.findByText('Download');
    fireEvent.click(downloadButton);

    await waitFor(() => {
      expect(api.downloadItem).toHaveBeenCalledWith('drive1', 'file1', 'document.pdf');
    });
  });

  it('navigates back to sites', async () => {
    api.getStatus.mockResolvedValue({ connected: true });
    api.listSites.mockResolvedValue({
      value: [{ id: 'site1', displayName: 'Site 1' }],
    });
    api.listDrives.mockResolvedValue({
      value: [{ id: 'drive1', name: 'Documents' }],
    });

    render(<FileBrowser />);

    const siteButton = await screen.findByText('Site 1');
    fireEvent.click(siteButton);

    const backButton = await screen.findByText('← Back to Sites');
    fireEvent.click(backButton);

    await waitFor(() => {
      expect(screen.getByText('SharePoint Sites')).toBeInTheDocument();
    });
  });
});
