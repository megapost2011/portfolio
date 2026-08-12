using System.Collections.Generic;
using System.Windows.Forms;

namespace BusinessManagementSuite.UI;

public sealed class PageHost : Panel
{
    private readonly Dictionary<string, UserControl> pages = new();

    private UserControl? currentPage;

    public PageHost()
    {
        Dock = DockStyle.Fill;

        BackColor = Theme.Surface;
    }

    //----------------------------------------------------
    // ページ登録
    //----------------------------------------------------

    public void RegisterPage(
        string key,
        UserControl page)
    {
        if (pages.ContainsKey(key))
            return;

        page.Dock = DockStyle.Fill;

        page.Visible = false;

        pages.Add(key, page);

        Controls.Add(page);
    }

    //----------------------------------------------------
    // ページ切替
    //----------------------------------------------------

    public void ShowPage(string key)
    {
        if (!pages.TryGetValue(key, out UserControl? page))
            return;

        if (currentPage != null)
            currentPage.Visible = false;

        currentPage = page;

        currentPage.Visible = true;

        currentPage.BringToFront();
    }

    //----------------------------------------------------
    // 現在ページ
    //----------------------------------------------------

    public string? CurrentPageKey
    {
        get
        {
            foreach (var pair in pages)
            {
                if (pair.Value == currentPage)
                    return pair.Key;
            }

            return null;
        }
    }

    //----------------------------------------------------
    // ページ一覧
    //----------------------------------------------------

    public IReadOnlyDictionary<string, UserControl> Pages
    {
        get
        {
            return pages;
        }
    }

    //----------------------------------------------------
    // 登録確認
    //----------------------------------------------------

    public bool Contains(string key)
    {
        return pages.ContainsKey(key);
    }

    //----------------------------------------------------
    // 削除
    //----------------------------------------------------

    public void RemovePage(string key)
    {
        if (!pages.TryGetValue(key, out UserControl? page))
            return;

        Controls.Remove(page);

        pages.Remove(key);

        page.Dispose();

        if (currentPage == page)
            currentPage = null;
    }

    //----------------------------------------------------
    // 全削除
    //----------------------------------------------------

    public void ClearPages()
    {
        foreach (var page in pages.Values)
        {
            page.Dispose();
        }

        pages.Clear();

        Controls.Clear();

        currentPage = null;
    }
}