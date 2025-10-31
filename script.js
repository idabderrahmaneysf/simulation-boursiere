// Fonction de calcul des frais (identique à la version VBA)
function calculerFrais(montantBrut, estAchat = true) {
    const commission = montantBrut * 0.6 / 100;    // 0.6% de commission
    const impot = montantBrut * 0.1 / 100;         // 0.1% d'impôt de bourse
    const tvaMontant = (commission + impot) * 10 / 100;  // 10% de TVA
    const totalFrais = commission + impot + tvaMontant;
    
    return totalFrais;
}

// Fonction de calcul du gain net
function calculerGainNet() {
    // Validation des données
    if (!validerDonnees()) {
        return;
    }

    // Récupération des valeurs
    const valeur = document.getElementById('valeur').value;
    const quantite = parseInt(document.getElementById('quantite').value);
    const coursAchat = parseFloat(document.getElementById('coursAchat').value);
    const coursVente = parseFloat(document.getElementById('coursVente').value);

    // Calculs détaillés
    const montantBrutAchat = quantite * coursAchat;
    const fraisAchat = calculerFrais(montantBrutAchat, true);
    const montantNetAchat = montantBrutAchat + fraisAchat;

    const montantBrutVente = quantite * coursVente;
    const fraisVente = calculerFrais(montantBrutVente, false);
    const montantNetVente = montantBrutVente - fraisVente;

    const gainBrut = montantNetVente - montantNetAchat;
    const gainNet = gainBrut * (1 - 0.15); // 15% de frais d'état

    // Affichage du résultat
    afficherResultat(gainNet, valeur, {
        montantBrutAchat,
        fraisAchat,
        montantNetAchat,
        montantBrutVente,
        fraisVente,
        montantNetVente,
        gainBrut,
        gainNet
    });
}

// Fonction de validation des données
function validerDonnees() {
    const valeur = document.getElementById('valeur').value;
    const quantite = document.getElementById('quantite').value;
    const coursAchat = document.getElementById('coursAchat').value;
    const coursVente = document.getElementById('coursVente').value;

    if (valeur.trim() === '') {
        alert('Veuillez saisir le nom de la valeur.');
        document.getElementById('valeur').focus();
        return false;
    }

    if (quantite === '' || isNaN(quantite) || parseInt(quantite) <= 0) {
        alert('Veuillez saisir une quantité valide (supérieure à zéro).');
        document.getElementById('quantite').focus();
        return false;
    }

    if (coursAchat === '' || isNaN(coursAchat) || parseFloat(coursAchat) <= 0) {
        alert('Veuillez saisir un cours d\'achat valide (supérieur à zéro).');
        document.getElementById('coursAchat').focus();
        return false;
    }

    if (coursVente === '' || isNaN(coursVente) || parseFloat(coursVente) <= 0) {
        alert('Veuillez saisir un cours de vente valide (supérieur à zéro).');
        document.getElementById('coursVente').focus();
        return false;
    }

    return true;
}

// Fonction d'affichage du résultat
function afficherResultat(gainNet, valeur, details) {
    const resultatDiv = document.getElementById('resultat');
    const detailsDiv = document.getElementById('detailsCalcul');
    const detailsContent = document.getElementById('detailsContent');
    
    let classeCss = 'gain-neutre';
    let icone = '⚖️';
    
    if (gainNet > 0) {
        classeCss = 'gain-positif';
        icone = '📈';
    } else if (gainNet < 0) {
        classeCss = 'gain-negatif';
        icone = '📉';
    }

    const gainFormate = new Intl.NumberFormat('fr-FR', { 
        minimumFractionDigits: 2, 
        maximumFractionDigits: 2 
    }).format(gainNet);

    resultatDiv.innerHTML = `
        ${icone} <strong>${valeur}</strong><br>
        Gain NET: <strong>${gainFormate} MAD</strong>
        <br><small style="cursor: pointer; color: #007cba;" onclick="afficherDetails()">Voir les détails du calcul</small>
    `;
    resultatDiv.className = `result ${classeCss}`;
    resultatDiv.classList.remove('hidden');

    // Préparer les détails
    detailsContent.innerHTML = `
        <table style="width: 100%; border-collapse: collapse;">
            <tr><td><strong>Montant brut d'achat:</strong></td><td>${formatMontant(details.montantBrutAchat)} MAD</td></tr>
            <tr><td><strong>Frais d'achat:</strong></td><td>${formatMontant(details.fraisAchat)} MAD</td></tr>
            <tr><td><strong>Montant net d'achat:</strong></td><td>${formatMontant(details.montantNetAchat)} MAD</td></tr>
            <tr><td style="border-bottom: 1px solid #ddd; padding-bottom: 5px;"></td><td style="border-bottom: 1px solid #ddd; padding-bottom: 5px;"></td></tr>
            <tr><td><strong>Montant brut de vente:</strong></td><td>${formatMontant(details.montantBrutVente)} MAD</td></tr>
            <tr><td><strong>Frais de vente:</strong></td><td>${formatMontant(details.fraisVente)} MAD</td></tr>
            <tr><td><strong>Montant net de vente:</strong></td><td>${formatMontant(details.montantNetVente)} MAD</td></tr>
            <tr><td style="border-bottom: 1px solid #ddd; padding-bottom: 5px;"></td><td style="border-bottom: 1px solid #ddd; padding-bottom: 5px;"></td></tr>
            <tr><td><strong>Gain brut:</strong></td><td>${formatMontant(details.gainBrut)} MAD</td></tr>
            <tr><td><strong>Frais d'état (15%):</strong></td><td>${formatMontant(details.gainBrut * 0.15)} MAD</td></tr>
            <tr><td style="border-top: 2px solid #333;"><strong>GAIN NET:</strong></td><td style="border-top: 2px solid #333;"><strong>${formatMontant(details.gainNet)} MAD</strong></td></tr>
        </table>
    `;
}

function formatMontant(montant) {
    return new Intl.NumberFormat('fr-FR', { 
        minimumFractionDigits: 2, 
        maximumFractionDigits: 2 
    }).format(montant);
}

function afficherDetails() {
    const detailsDiv = document.getElementById('detailsCalcul');
    detailsDiv.style.display = detailsDiv.style.display === 'none' ? 'block' : 'none';
}

function nouvelleSimulation() {
    document.getElementById('valeur').value = '';
    document.getElementById('quantite').value = '';
    document.getElementById('coursAchat').value = '';
    document.getElementById('coursVente').value = '';
    document.getElementById('resultat').classList.add('hidden');
    document.getElementById('detailsCalcul').style.display = 'none';
    document.getElementById('valeur').focus();
}

function exporterExcel() {
    if (!validerDonnees()) {
        return;
    }

    const valeur = document.getElementById('valeur').value;
    const quantite = document.getElementById('quantite').value;
    const coursAchat = document.getElementById('coursAchat').value;
    const coursVente = document.getElementById('coursVente').value;

    // Calcul pour obtenir les données
    const quantiteNum = parseInt(quantite);
    const coursAchatNum = parseFloat(coursAchat);
    const coursVenteNum = parseFloat(coursVente);
    
    const montantBrutAchat = quantiteNum * coursAchatNum;
    const fraisAchat = calculerFrais(montantBrutAchat, true);
    const montantNetAchat = montantBrutAchat + fraisAchat;
    const montantBrutVente = quantiteNum * coursVenteNum;
    const fraisVente = calculerFrais(montantBrutVente, false);
    const montantNetVente = montantBrutVente - fraisVente;
    const gainBrut = montantNetVente - montantNetAchat;
    const gainNet = gainBrut * (1 - 0.15);

    // Création du contenu CSV
    const csvContent = [
        ['Simulation Boursière - Export Excel'],
        ['Valeur', 'Quantité', 'Cours Achat (MAD)', 'Cours Vente (MAD)', 'Gain NET (MAD)'],
        [valeur, quantite, coursAchat, coursVente, gainNet.toFixed(2)],
        [],
        ['DÉTAILS DU CALCUL'],
        ['Montant brut achat:', `${montantBrutAchat.toFixed(2)} MAD`],
        ['Frais achat:', `${fraisAchat.toFixed(2)} MAD`],
        ['Montant net achat:', `${montantNetAchat.toFixed(2)} MAD`],
        ['Montant brut vente:', `${montantBrutVente.toFixed(2)} MAD`],
        ['Frais vente:', `${fraisVente.toFixed(2)} MAD`],
        ['Montant net vente:', `${montantNetVente.toFixed(2)} MAD`],
        ['Gain brut:', `${gainBrut.toFixed(2)} MAD`],
        ['Frais d\'état (15%):', `${(gainBrut * 0.15).toFixed(2)} MAD`],
        ['GAIN NET:', `${gainNet.toFixed(2)} MAD`]
    ].map(row => row.join(',')).join('\n');

    // Téléchargement du fichier
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', `simulation_${valeur}_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Calcul en temps réel
document.addEventListener('DOMContentLoaded', function() {
    const inputs = ['quantite', 'coursAchat', 'coursVente'];
    inputs.forEach(id => {
        document.getElementById(id).addEventListener('input', function() {
            if (validerDonnees()) {
                calculerGainNet();
            }
        });
    });
});