import { describe, it, expect, beforeEach, vi } from 'vitest';
import api from '../api';

// Mock fetch
global.fetch = vi.fn();

describe('API Client', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('getConnectUrl returns auth URL and state', async () => {
    global.fetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        auth_url: 'https://login.microsoftonline.com/...',
        state: 'state123',
      }),
    });

    const result = await api.getConnectUrl();

    expect(result.auth_url).toBeTruthy();
    expect(result.state).toBeTruthy();
  });

  it('getStatus returns connection status', async () => {
    global.fetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        connected: true,
        microsoft_user_id: 'user@example.com',
      }),
    });

    const result = await api.getStatus();

    expect(result.connected).toBe(true);
    expect(result.microsoft_user_id).toBeTruthy();
  });

  it('listSites returns sites array', async () => {
    global.fetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        value: [
          { id: 'site1', displayName: 'Site 1' },
        ],
      }),
    });

    const result = await api.listSites();

    expect(result.value).toHaveLength(1);
    expect(result.value[0].id).toBe('site1');
  });

  it('listSites includes search parameter', async () => {
    global.fetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({ value: [] }),
    });

    await api.listSites('Marketing');

    const call = global.fetch.mock.calls[0][0];
    expect(call.includes('search=Marketing')).toBe(true);
  });

  it('listDrives returns drives array', async () => {
    global.fetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        value: [
          { id: 'drive1', name: 'Documents' },
        ],
      }),
    });

    const result = await api.listDrives('site123');

    expect(result.value).toHaveLength(1);
    expect(result.value[0].id).toBe('drive1');
  });

  it('listItems returns items array', async () => {
    global.fetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        value: [
          { id: 'item1', name: 'file.txt' },
        ],
      }),
    });

    const result = await api.listItems('drive123');

    expect(result.value).toHaveLength(1);
  });

  it('listItems includes item_id parameter', async () => {
    global.fetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({ value: [] }),
    });

    await api.listItems('drive123', 'folder123');

    const call = global.fetch.mock.calls[0][0];
    expect(call.includes('item_id=folder123')).toBe(true);
  });

  it('downloadItem throws on API error', async () => {
    global.fetch.mockResolvedValueOnce({
      ok: false,
    });

    await expect(
      api.downloadItem('drive123', 'item123', 'file.txt')
    ).rejects.toThrow('Failed to download item');
  });
});
