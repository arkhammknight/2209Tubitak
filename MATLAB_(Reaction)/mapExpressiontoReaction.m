% Model ve veri yükleme
model = readCbModel('C:\Users\omer\cobratoolbox\Human-GEM.mat');
expressdata = readcell('C:\Users\omer\TCGADATA\TCGASTAD\fpkmstad.xlsx');

% RNA-seq verisinden gen isimleri ve ekspresyon matrisini ayır
geneSymbols = expressdata(2:end, 1);                       % Sadece 1. sütundaki gen isimlerini alıyor
expressionMatrix = cell2mat(expressdata(2:end, 2:end));    % 2. sütundan itibaren kalan verileri sayısal (double) matrise çeviriyor.

% model.geneShortNames ile RNA-seq verisindeki gen sembollerini eşleştir
[commonGenes, idxRNA, idxModel] = intersect(geneSymbols, model.geneShortNames, 'stable');
% intersect fonksiyonu iki listeyi karşılaştırır ve ortak olan genleri bulur:
% commonGenes: Ortak bulunan gen isimlerinin listesi.
% idxRNA: RNA-Seq verisindeki geneSymbols içindeki ortak genlerin indeksleri.
% idxModel: Modeldeki model.geneShortNames içindeki ortak genlerin indeksleri.
% 'stable' opsiyonu: RNA-seq verisindeki sıralamayı korur.

% matchedExpression: model ile eşleşen genlerin ekspresyonları
matchedExpression = expressionMatrix(idxRNA, :);

% ExpressionData yapısını oluştur (eşleşen genler üzerinden)
expressionData.gene = model.genes(idxModel);
expressionData.value = matchedExpression;         % eşleşmiş ekspresyon verisi

% Reaksiyon ekspresyonlarını her sütun için hesapla
numSamples = size(expressionData.value, 2); % Örnek sayısı (294)
rxnExpr = zeros(length(model.rxns), numSamples); % 12971 x 294 matris
% rxnExpr: Reaksiyonlar x Örnekler büyüklüğünde bir boş matris oluşturuluyor.
% Her reaksiyon için her örnek bazında ekspresyon hesaplanıyor.

% Her Örnek İçin mapExpressionToReactions Kullanımı
for i = 1:numSamples
    % Her sütun için expressionData oluştur
    tempExpressionData.gene = expressionData.gene; 
    tempExpressionData.value = expressionData.value(:, i); % i'nci sütun % Her seferinde tek bir örneğe ait ekspresyon verisini alıyoruz.
    
    % mapExpressionToReactions ile reaksiyon ekspresyonlarını hesapla
    [tempRxnExpr, parsedGPR] = mapExpressionToReactions(model, tempExpressionData); % tempRxnExpr: Bir vektör (12971 x 1), her reaksiyon için o örnekteki ekspresyon değeri.
    
    % Sonuçları rxnExpr matrisine ekle
    rxnExpr(:, i) = tempRxnExpr;
end

% rxnExpr boyutlarını kontrol et
disp(size(rxnExpr)); % 12971 x 294 olmalı

writematrix(rxnExpr, 'C:\Users\omer\TCGADATA\TCGASTAD\rxnExprstad.xlsx');
