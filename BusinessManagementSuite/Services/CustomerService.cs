using System;
using System.Collections.Generic;

using BusinessManagementSuite.Data;
using BusinessManagementSuite.Models;

namespace BusinessManagementSuite.Services;

public sealed class CustomerService
{
    private readonly CustomerRepository _repo = new();

    //---------------------------------------------------------
    // Query
    //---------------------------------------------------------

    public List<Customer> GetAll()              => _repo.GetAll();
    public List<Customer> GetActiveCustomers()  => _repo.GetActiveCustomers();
    public Customer?      GetById(long id)      => _repo.GetById(id);
    public Customer?      GetByCode(string code)=> _repo.GetByCode(code);
    public List<Customer> Search(string keyword)=> _repo.Search(keyword);
    public bool           Exists(long id)       => _repo.Exists(id);
    public bool           CodeExists(string c)  => _repo.CodeExists(c);
    public int            Count()               => _repo.Count();
    public int            ActiveCount()         => _repo.ActiveCount();
    public int            InactiveCount()       => _repo.InactiveCount();

    //---------------------------------------------------------
    // Create (factory)
    //---------------------------------------------------------

    public Customer Create()
    {
        return new Customer
        {
            CustomerCode = _repo.GenerateCustomerCode(),
            IsActive     = true
        };
    }

    //---------------------------------------------------------
    // Save
    //---------------------------------------------------------

    public void Save(Customer customer)
    {
        Validate(customer);

        Customer? existing = _repo.GetByCode(customer.CustomerCode);

        if (existing != null && existing.Id != customer.Id)
        {
            throw new InvalidOperationException(
                $"Customer code '{customer.CustomerCode}' already exists.");
        }

        _repo.Save(customer);
    }

    //---------------------------------------------------------
    // Delete / Restore
    //---------------------------------------------------------

    public bool Delete(long id)
    {
        if (!_repo.Exists(id)) return false;
        return _repo.Delete(id);
    }

    public bool Restore(long id)        => _repo.Restore(id);
    public bool DeletePhysical(long id) => _repo.DeletePhysical(id);

    //---------------------------------------------------------
    // Validation
    //---------------------------------------------------------

    public void Validate(Customer customer)
    {
        if (customer == null)
            throw new ArgumentNullException(nameof(customer));

        if (string.IsNullOrWhiteSpace(customer.CustomerCode))
            throw new ArgumentException("Customer code is required.");

        if (string.IsNullOrWhiteSpace(customer.CustomerName))
            throw new ArgumentException("Customer name is required.");

        ValidateEmail(customer.Email);
    }

    public void ValidateEmail(string? email)
    {
        if (string.IsNullOrWhiteSpace(email)) return;

        try
        {
            _ = new System.Net.Mail.MailAddress(email);
        }
        catch
        {
            throw new ArgumentException("Invalid email address.");
        }
    }

    //---------------------------------------------------------
    // Export / Import (stub – Phase 5)
    //---------------------------------------------------------

    public void ExportCsv(string fileName)
        => throw new NotImplementedException("CSV export: Phase 5.");

    public void ExportExcel(string fileName)
        => throw new NotImplementedException("Excel export: Phase 5.");

    public void ImportCsv(string fileName)
        => throw new NotImplementedException("CSV import: Phase 5.");

    //---------------------------------------------------------
    // Statistics
    //---------------------------------------------------------

    public CustomerStatistics GetStatistics()
    {
        return new CustomerStatistics
        {
            TotalCustomers    = _repo.Count(),
            ActiveCustomers   = _repo.ActiveCount(),
            InactiveCustomers = _repo.InactiveCount()
        };
    }
}
