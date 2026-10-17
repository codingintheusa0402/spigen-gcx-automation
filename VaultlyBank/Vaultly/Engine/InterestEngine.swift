import Foundation

/// Moves time forward for every product account:
///  - deposit / savings: daily simple interest accrues into `accruedInterest`,
///    paid out (after tax) at maturity.
///  - savings: monthly installment auto-transferred from the linked checking account.
///  - loan: monthly equal-payment (원리금균등) repayment pulled from the linked account;
///    interest for the month is added to the debt first, then the payment reduces it.
enum InterestEngine {
    static let daysPerYear = 365.0
    static let calendar = Calendar.current

    /// Equal monthly payment for a loan.
    static func annuityPayment(principal: Decimal, annualRate: Double, months: Int) -> Decimal {
        guard months > 0 else { return principal }
        let p = principal.doubleValue
        let r = annualRate / 12
        if r == 0 { return (principal / Decimal(months)).rounded() }
        let pay = p * r / (1 - pow(1 + r, Double(-months)))
        return Decimal(double: pay).rounded()
    }

    /// Projected payout (before tax) at maturity for savings/deposit.
    static func projectedMaturity(terms: ProductTerms, kind: AccountKind) -> (principal: Decimal, interest: Decimal) {
        let r = terms.annualRate
        switch kind {
        case .deposit:
            let interest = terms.principal.doubleValue * r * Double(terms.termMonths) / 12
            return (terms.principal, Decimal(double: interest).rounded())
        case .installmentSavings:
            // Korean 적금 simple interest: k-th installment earns rate * (n-k+1)/12
            let n = terms.termMonths
            let m = terms.monthlyAmount.doubleValue
            let interest = (1...max(n, 1)).reduce(0.0) { $0 + m * r * Double(n - $1 + 1) / 12 }
            return (terms.monthlyAmount * Decimal(n), Decimal(double: interest).rounded())
        default:
            return (0, 0)
        }
    }

    static func accrue(now: Date = Date(), data: inout BankData, taxRate: Double) {
        for idx in data.accounts.indices {
            guard var terms = data.accounts[idx].terms, !terms.isClosed else { continue }
            let account = data.accounts[idx]
            switch account.kind {
            case .deposit:
                accrueDaily(now: now, idx: idx, terms: &terms, data: &data)
                settleIfMatured(now: now, idx: idx, terms: &terms, data: &data, taxRate: taxRate)
            case .installmentSavings:
                processInstallments(now: now, idx: idx, terms: &terms, data: &data)
                accrueDaily(now: now, idx: idx, terms: &terms, data: &data)
                settleIfMatured(now: now, idx: idx, terms: &terms, data: &data, taxRate: taxRate)
            case .loan:
                processRepayments(now: now, idx: idx, terms: &terms, data: &data)
            case .checking:
                break
            }
            data.accounts[idx].terms = terms
        }
    }

    // MARK: Daily interest

    private static func accrueDaily(now: Date, idx: Int, terms: inout ProductTerms, data: inout BankData) {
        let end = min(now, terms.maturityDate)
        let days = calendar.dateComponents([.day], from: calendar.startOfDay(for: terms.lastAccrualDate),
                                           to: calendar.startOfDay(for: end)).day ?? 0
        guard days > 0 else { return }
        let balance = data.accounts[idx].balance.doubleValue
        let interest = balance * terms.annualRate / daysPerYear * Double(days)
        terms.accruedInterest += Decimal(double: interest)
        terms.lastAccrualDate = end
    }

    private static func settleIfMatured(now: Date, idx: Int, terms: inout ProductTerms, data: inout BankData, taxRate: Double) {
        guard now >= terms.maturityDate else { return }
        let gross = terms.accruedInterest.rounded()
        let tax = (gross * Decimal(double: taxRate)).rounded()
        let net = gross - tax
        let id = data.accounts[idx].id
        if net != 0 {
            Ledger.apply(amount: net, to: id, date: terms.maturityDate, counterparty: data.settings.brand.bank,
                         memo: "만기 이자 (세전 \(Money.format(gross, code: data.settings.currencyCode)))",
                         category: .maturity, data: &data)
        }
        terms.accruedInterest = 0
        terms.isClosed = true
        // Move the matured amount to the linked account so it shows up as spendable money.
        if let linked = terms.linkedAccountID, data.accounts.contains(where: { $0.id == linked }) {
            let total = data.accounts[idx].balance
            if total > 0 {
                Ledger.apply(amount: -total, to: id, date: terms.maturityDate, counterparty: "만기 해지",
                             memo: "만기 지급", category: .maturity, data: &data)
                Ledger.apply(amount: total, to: linked, date: terms.maturityDate,
                             counterparty: data.accounts[idx].name, memo: "만기 지급", category: .maturity, data: &data)
            }
        }
    }

    // MARK: Monthly installments (적금)

    private static func processInstallments(now: Date, idx: Int, terms: inout ProductTerms, data: inout BankData) {
        var guardCount = 0
        while let due = calendar.date(byAdding: .month, value: 1, to: terms.lastInstallmentDate),
              due <= now, due <= terms.maturityDate, guardCount < 2400 {
            guardCount += 1
            // accrue interest up to the installment date before the balance changes
            let savedNow = due
            accrueDaily(now: savedNow, idx: idx, terms: &terms, data: &data)
            let id = data.accounts[idx].id
            if let linked = terms.linkedAccountID,
               let li = data.accounts.firstIndex(where: { $0.id == linked }) {
                if data.accounts[li].balance >= terms.monthlyAmount {
                    Ledger.apply(amount: -terms.monthlyAmount, to: linked, date: due, counterparty: data.accounts[idx].name,
                                 memo: "적금 자동이체", category: .autoTransfer, data: &data)
                    Ledger.apply(amount: terms.monthlyAmount, to: id, date: due, counterparty: data.accounts[li].name,
                                 memo: "자동이체 입금", category: .autoTransfer, data: &data)
                } else {
                    Ledger.apply(amount: 0, to: id, date: due, counterparty: data.accounts[li].name,
                                 memo: "자동이체 실패 (잔액 부족)", category: .autoTransfer, data: &data)
                }
            } else {
                Ledger.apply(amount: terms.monthlyAmount, to: id, date: due, counterparty: "자동이체",
                             memo: "월 납입", category: .autoTransfer, data: &data)
            }
            terms.lastInstallmentDate = due
        }
    }

    // MARK: Monthly repayments (대출)

    private static func processRepayments(now: Date, idx: Int, terms: inout ProductTerms, data: inout BankData) {
        var guardCount = 0
        while let due = calendar.date(byAdding: .month, value: 1, to: terms.lastInstallmentDate),
              due <= now, guardCount < 2400 {
            guardCount += 1
            let id = data.accounts[idx].id
            let debt = data.accounts[idx].balance
            guard debt > 0 else { terms.isClosed = true; break }
            let interest = (debt * Decimal(double: terms.annualRate / 12)).rounded()
            Ledger.apply(amount: interest, to: id, date: due, counterparty: data.settings.brand.bank,
                         memo: "대출 이자", category: .interest, data: &data)
            let payment = min(terms.monthlyAmount, debt + interest)
            if let linked = terms.linkedAccountID,
               let li = data.accounts.firstIndex(where: { $0.id == linked }) {
                if data.accounts[li].balance >= payment {
                    Ledger.apply(amount: -payment, to: linked, date: due, counterparty: data.accounts[idx].name,
                                 memo: "대출 원리금 상환", category: .repayment, data: &data)
                    Ledger.apply(amount: -payment, to: id, date: due, counterparty: data.accounts[li].name,
                                 memo: "원리금 상환 (이자 \(Money.format(interest, code: data.settings.currencyCode)))",
                                 category: .repayment, data: &data)
                } else {
                    Ledger.apply(amount: 0, to: id, date: due, counterparty: data.accounts[li].name,
                                 memo: "상환 실패 (잔액 부족) — 연체", category: .repayment, data: &data)
                }
            } else {
                Ledger.apply(amount: -payment, to: id, date: due, counterparty: "자동상환",
                             memo: "원리금 상환", category: .repayment, data: &data)
            }
            terms.lastInstallmentDate = due
            if data.accounts[idx].balance <= 0 {
                data.accounts[idx].balance = 0
                terms.isClosed = true
                break
            }
        }
    }
}
